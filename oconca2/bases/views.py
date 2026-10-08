import os
import json
import tempfile
from io import BytesIO
from datetime import datetime, timedelta
import jwt
import openpyxl
import docx

from django.shortcuts import render
from django.http import JsonResponse, FileResponse, HttpResponse, HttpResponseNotFound, HttpResponseBadRequest, HttpResponseForbidden
from django.views.decorators.csrf import csrf_exempt
from django.contrib.auth import authenticate
from django.contrib.auth.models import User
from django.conf import settings

from .models import Evaluacion

SECRET_KEY = getattr(settings, 'SECRET_KEY', 'tu_clave_secreta_super_segura')
ALGORITHM = "HS256"

# Helper para token JWT
def generate_access_token(user):
    payload = {
        "sub": user.username,
        "exp": datetime.utcnow() + timedelta(days=7)
    }
    return jwt.encode(payload, SECRET_KEY, algorithm=ALGORITHM)

def get_current_user(request):
    auth_header = request.headers.get('Authorization', '')
    if not auth_header.startswith('Bearer '):
        return None
    token = auth_header.split(' ')[1]
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        username = payload.get("sub")
        if username:
            return User.objects.filter(username=username).first()
    except Exception:
        return None
    return None

def index(request):
    """Vista principal que carga el dashboard en AdminLTE 3 con datos CIMTRA"""
    base_dir = settings.BASE_DIR
    json_path = os.path.join(base_dir, "cimtra_data.json")
    cimtra_data_json = "[]"
    if os.path.exists(json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            cimtra_data_json = f.read()
            
    context = {
        "cimtra_data_json": cimtra_data_json
    }
    return render(request, 'bases/index.html', context)


@csrf_exempt
def api_register(request):
    if request.method != 'POST':
        return HttpResponseBadRequest("Método no permitido")
    
    try:
        data = json.loads(request.body.decode('utf-8'))
        username = (data.get('username') or '').strip()
        password = data.get('password')
    except Exception:
        return JsonResponse({"detail": "Datos inválidos"}, status=400)
    
    if not username or not password:
        return JsonResponse({"detail": "Usuario y contraseña requeridos"}, status=400)
        
    if User.objects.filter(username__iexact=username).exists():
        return JsonResponse({"detail": "El nombre de usuario ya está registrado"}, status=400)
        
    User.objects.create_user(username=username, password=password)
    return JsonResponse({"message": "Usuario registrado exitosamente"})

@csrf_exempt
def api_token(request):
    """Soporta tanto Form Data (OAuth2 standard) como JSON login"""
    if request.method != 'POST':
        return HttpResponseBadRequest("Método no permitido")
    
    username = None
    password = None

    if request.body:
        try:
            data = json.loads(request.body.decode('utf-8'))
            if isinstance(data, dict):
                username = data.get('username')
                password = data.get('password')
        except Exception:
            pass

    if not username:
        username = request.POST.get('username')
    if not password:
        password = request.POST.get('password')

    if not username or not password:
        return JsonResponse({"detail": "Usuario y contraseña requeridos"}, status=400)

    username = username.strip()

    # Autenticar considerando coincidencia insensible a mayúsculas/minúsculas
    user_match = User.objects.filter(username__iexact=username).first()
    if user_match:
        user = authenticate(username=user_match.username, password=password)
    else:
        user = authenticate(username=username, password=password)

    if not user:
        return JsonResponse({"detail": "Usuario o contraseña incorrectos"}, status=401)

    token = generate_access_token(user)
    return JsonResponse({"access_token": token, "token_type": "bearer"})

def api_me(request):
    user = get_current_user(request)
    if not user:
        return JsonResponse({"detail": "No autenticado"}, status=401)
    
    es_admin = user.is_staff or user.is_superuser
    return JsonResponse({"username": user.username, "es_admin": es_admin})

@csrf_exempt
def api_guardar(request):
    if request.method != 'POST':
        return HttpResponseBadRequest("Método no permitido")
        
    user = get_current_user(request)
    if not user:
        return JsonResponse({"detail": "No autenticado"}, status=401)

    try:
        payload = json.loads(request.body.decode('utf-8'))
    except Exception:
        return JsonResponse({"detail": "Payload inválido"}, status=400)

    entidad = payload.get("entidad")
    periodo = payload.get("periodo")
    tipo = payload.get("tipo", "Capital")
    cuestionario_data = payload.get("cuestionario_data", [])

    if not entidad or not periodo:
        return JsonResponse({"detail": "Entidad y período son requeridos"}, status=400)

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    evaluacion = Evaluacion.objects.filter(entidad=entidad, periodo=periodo, tipo=tipo).first()

    if evaluacion:
        evaluacion.datos_cuestionario = json.dumps(cuestionario_data)
        evaluacion.ultima_modificacion = now_str
        evaluacion.usuario = user
        evaluacion.save()
    else:
        Evaluacion.objects.create(
            usuario=user,
            entidad=entidad,
            periodo=periodo,
            tipo=tipo,
            datos_cuestionario=json.dumps(cuestionario_data),
            fecha_creacion=now_str,
            ultima_modificacion=now_str
        )

    return JsonResponse({"message": "Progreso guardado correctamente"})

def api_cargar(request):
    user = get_current_user(request)
    if not user:
        return JsonResponse({"detail": "No autenticado"}, status=401)

    entidad = request.GET.get("entidad")
    periodo = request.GET.get("periodo")
    tipo = request.GET.get("tipo", "Capital")

    evaluacion = Evaluacion.objects.filter(entidad=entidad, periodo=periodo, tipo=tipo).first()

    if not evaluacion or not evaluacion.datos_cuestionario:
        return JsonResponse({"detail": "No se encontró registro para esta entidad y período"}, status=404)

    return JsonResponse({
        "entidad": evaluacion.entidad,
        "periodo": evaluacion.periodo,
        "tipo": evaluacion.tipo,
        "cuestionario_data": json.loads(evaluacion.datos_cuestionario),
        "ultima_modificacion": evaluacion.ultima_modificacion or "",
        "ultimo_usuario": evaluacion.usuario.username if evaluacion.usuario else "Desconocido"
    })

def _calcular_stats_evaluacion(datos_json):
    """Calcula indicadores cuantitativos a partir del cuestionario JSON"""
    try:
        campos = json.loads(datos_json) if isinstance(datos_json, str) else datos_json
    except Exception:
        return {
            "puntos_obtenidos": 0, "puntos_maximos": 0, "porcentaje": 0,
            "criterios_si": 0, "criterios_no": 0, "criterios_na": 0,
            "total_criterios": 0, "total_observaciones": 0,
            "deficiencias": {"no_existe": 0, "no_actualizada": 0, "ilegible": 0, "enlace_roto": 0},
            "desglose_campos": []
        }

    puntos_obtenidos = 0
    puntos_maximos = 0
    criterios_si = 0
    criterios_no = 0
    criterios_na = 0
    total_observaciones = 0
    deficiencias = {"no_existe": 0, "no_actualizada": 0, "ilegible": 0, "enlace_roto": 0}
    desglose_campos = []

    for campo in campos:
        c_puntos_obtenidos = 0
        c_puntos_maximos = 0
        c_si = 0
        c_no = 0
        c_na = 0
        c_obs = 0

        for bloque in campo.get('bloques', []):
            for aspecto in bloque.get('aspectos', []):
                for c in aspecto.get('criterios', []):
                    try:
                        pts = int(c.get('puntos', 1))
                    except Exception:
                        pts = 1
                    puntos_maximos += pts
                    c_puntos_maximos += pts

                    cumple = c.get('cumple')
                    if cumple == 'si':
                        criterios_si += 1
                        puntos_obtenidos += pts
                        c_si += 1
                        c_puntos_obtenidos += pts
                    elif cumple == 'no':
                        criterios_no += 1
                        c_no += 1
                    else:
                        criterios_na += 1
                        c_na += 1

                    comentario = (c.get('comentario') or '').strip()
                    if comentario:
                        total_observaciones += 1
                        c_obs += 1

                    errs = c.get('errores', {})
                    for k in deficiencias.keys():
                        if errs.get(k):
                            deficiencias[k] += 1

        c_porcentaje = round((c_puntos_obtenidos / c_puntos_maximos * 100), 2) if c_puntos_maximos > 0 else 0
        desglose_campos.append({
            "id": campo.get("id"),
            "nombre": campo.get("nombre", ""),
            "puntos_obtenidos": c_puntos_obtenidos,
            "puntos_maximos": c_puntos_maximos,
            "porcentaje": c_porcentaje,
            "criterios_si": c_si,
            "criterios_no": c_no,
            "criterios_na": c_na,
            "total_criterios": c_si + c_no + c_na,
            "observaciones": c_obs
        })

    total_criterios = criterios_si + criterios_no + criterios_na
    porcentaje = round((puntos_obtenidos / puntos_maximos * 100), 2) if puntos_maximos > 0 else 0

    return {
        "puntos_obtenidos": puntos_obtenidos,
        "puntos_maximos": puntos_maximos,
        "porcentaje": porcentaje,
        "criterios_si": criterios_si,
        "criterios_no": criterios_no,
        "criterios_na": criterios_na,
        "total_criterios": total_criterios,
        "total_observaciones": total_observaciones,
        "deficiencias": deficiencias,
        "desglose_campos": desglose_campos
    }

def api_evaluaciones(request):
    user = get_current_user(request)
    if not user:
        return JsonResponse({"detail": "No autenticado"}, status=401)

    evaluaciones = Evaluacion.objects.all().order_by('-id')
    resultado = []
    for ev in evaluaciones:
        stats = _calcular_stats_evaluacion(ev.datos_cuestionario)
        resultado.append({
            "id": ev.id,
            "entidad": ev.entidad,
            "periodo": ev.periodo,
            "tipo": ev.tipo,
            "ultima_modificacion": ev.ultima_modificacion or "",
            "ultimo_usuario": ev.usuario.username if ev.usuario else "Desconocido",
            "stats": stats
        })
    return JsonResponse(resultado, safe=False)

@csrf_exempt
def api_borrar(request):
    if request.method not in ['DELETE', 'POST']:
        return HttpResponseBadRequest("Método no permitido")
        
    user = get_current_user(request)
    if not user:
        return JsonResponse({"detail": "No autenticado"}, status=401)

    ev_id = request.GET.get("id")
    if not ev_id and request.body:
        try:
            body = json.loads(request.body.decode('utf-8'))
            ev_id = body.get("id")
        except Exception:
            pass

    if not ev_id:
        return JsonResponse({"detail": "ID requerido"}, status=400)

    evaluacion = Evaluacion.objects.filter(id=ev_id).first()
    if not evaluacion:
        return JsonResponse({"detail": "Evaluación no encontrada"}, status=404)

    evaluacion.delete()
    return JsonResponse({"message": "Evaluación eliminada correctamente"})

from .exporter import generar_excel, generar_word_plantilla, generar_word_observaciones
from .validator import validar_evaluacion_completa, validar_exportacion_excel, validar_exportacion_word, validar_datos_cuestionario

def api_exportar(request):
    user = get_current_user(request)
    if not user:
        return JsonResponse({"detail": "No autenticado"}, status=401)

    entidad = request.GET.get("entidad")
    periodo = request.GET.get("periodo")
    tipo = request.GET.get("tipo", "Capital")

    evaluacion = Evaluacion.objects.filter(entidad=entidad, periodo=periodo, tipo=tipo).first()
    if not evaluacion or not evaluacion.datos_cuestionario:
        return JsonResponse({"detail": "No se encontró registro para esta entidad y período"}, status=404)

    try:
        wb, temp_path = generar_excel(evaluacion)
        with open(temp_path, 'rb') as f:
            file_data = f.read()
    except Exception as e:
        return JsonResponse({"detail": f"Error al generar Excel: {str(e)}"}, status=500)
    finally:
        if 'temp_path' in locals() and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

    response = FileResponse(
        BytesIO(file_data),
        as_attachment=True,
        filename=f"CIMTRA_{entidad.replace(' ', '_')}_{periodo.replace(' ', '_')}.xlsx",
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return response

def api_exportar_word(request):
    user = get_current_user(request)
    if not user:
        return JsonResponse({"detail": "No autenticado"}, status=401)

    entidad = request.GET.get("entidad")
    periodo = request.GET.get("periodo")
    tipo = request.GET.get("tipo", "Capital")

    if tipo != "Congreso":
        return JsonResponse({"detail": "Formato Word solo disponible actualmente para Congreso"}, status=400)

    evaluacion = Evaluacion.objects.filter(entidad=entidad, periodo=periodo, tipo=tipo).first()
    if not evaluacion or not evaluacion.datos_cuestionario:
        return JsonResponse({"detail": "No se encontró registro para esta entidad y período"}, status=404)

    try:
        doc, temp_path = generar_word_plantilla(evaluacion)
        with open(temp_path, 'rb') as f:
            file_data = f.read()
    except Exception as e:
        return JsonResponse({"detail": f"Error al generar Word: {str(e)}"}, status=500)
    finally:
        if 'temp_path' in locals() and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

    response = FileResponse(
        BytesIO(file_data),
        as_attachment=True,
        filename=f"CIMTRA_{entidad.replace(' ', '_')}_{periodo.replace(' ', '_')}.docx",
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    return response

def api_exportar_observaciones_word(request):
    user = get_current_user(request)
    if not user:
        return JsonResponse({"detail": "No autenticado"}, status=401)

    entidad = request.GET.get("entidad")
    periodo = request.GET.get("periodo")
    tipo = request.GET.get("tipo", "Capital")

    evaluacion = Evaluacion.objects.filter(entidad=entidad, periodo=periodo, tipo=tipo).first()
    if not evaluacion or not evaluacion.datos_cuestionario:
        return JsonResponse({"detail": "No se encontró registro para esta entidad y período"}, status=404)

    try:
        doc, temp_path = generar_word_observaciones(evaluacion)
        with open(temp_path, 'rb') as f:
            file_data = f.read()
    except Exception as e:
        return JsonResponse({"detail": f"Error al generar informe Word: {str(e)}"}, status=500)
    finally:
        if 'temp_path' in locals() and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

    response = FileResponse(
        BytesIO(file_data),
        as_attachment=True,
        filename=f"Observaciones_CIMTRA_{entidad.replace(' ', '_')}_{periodo.replace(' ', '_')}.docx",
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    return response

def api_validar_exportacion(request):
    """Proceso interno para validar los datos, fórmulas y formatos de exportación de una evaluación"""
    user = get_current_user(request)
    if not user:
        return JsonResponse({"detail": "No autenticado"}, status=401)

    ev_id = request.GET.get("id")
    entidad = request.GET.get("entidad")
    periodo = request.GET.get("periodo")
    tipo = request.GET.get("tipo", "Capital")

    if ev_id:
        evaluacion = Evaluacion.objects.filter(id=ev_id).first()
    else:
        evaluacion = Evaluacion.objects.filter(entidad=entidad, periodo=periodo, tipo=tipo).first()

    if not evaluacion or not evaluacion.datos_cuestionario:
        return JsonResponse({"detail": "No se encontró registro de evaluación para validar"}, status=404)

    try:
        reporte = validar_evaluacion_completa(evaluacion)
        return JsonResponse(reporte)
    except Exception as e:
        return JsonResponse({"detail": f"Error durante la validación: {str(e)}"}, status=500)


