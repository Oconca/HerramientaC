import json
import os
from django.test import TestCase, Client
from django.contrib.auth.models import User
from django.core.management import call_command
from io import StringIO
from bases.models import Evaluacion
from bases.exporter import generar_excel, generar_word_plantilla, generar_word_observaciones
from bases.validator import (
    validar_datos_cuestionario,
    validar_exportacion_excel,
    validar_exportacion_word,
    validar_evaluacion_completa
)
from bases.views import generate_access_token

def crear_datos_prueba_capital():
    """Genera una estructura de datos de prueba para Capital con 9 bloques"""
    # Usamos la estructura básica de campos y bloques de CIMTRA Capital
    campos = [
        {
            "id": 1,
            "nombre": "Información a la Ciudadanía",
            "bloques": [
                {
                    "id": 1,
                    "nombre": "BLOQUE GASTOS",
                    "aspectos": [
                        {
                            "id": 1,
                            "titulo": "1. Gastos de comunicación",
                            "criterios": [
                                {
                                    "id": 1,
                                    "puntos": "1",
                                    "descripcion": "1.1 Información de gastos ejercidos y pagados en comunicación social.",
                                    "cumple": "si",
                                    "errores": {"no_existe": False, "no_actualizada": False, "no_corresponde": False, "ilegible": False, "enlace_roto": False},
                                    "comentario": ""
                                },
                                {
                                    "id": 2,
                                    "puntos": "1",
                                    "descripcion": "1.2 Comparativo directamente contra el monto ejercido y pagado del año inmediato anterior",
                                    "cumple": "no",
                                    "errores": {"no_existe": True, "no_actualizada": False, "no_corresponde": False, "ilegible": False, "enlace_roto": False},
                                    "comentario": "Falta información comparativa"
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    ]
    return json.dumps(campos)

def crear_datos_prueba_congreso():
    """Genera datos de prueba para Congreso"""
    campos = [
        {
            "id": 1,
            "nombre": "INFORMACION A LA CIUDADANIA",
            "bloques": [
                {
                    "id": 1,
                    "nombre": "INTEGRACIÓN Y ESTRUCTURA",
                    "aspectos": [
                        {
                            "id": 1,
                            "titulo": "1.- Conformación política",
                            "criterios": [
                                {
                                    "id": 1,
                                    "puntos": "1",
                                    "descripcion": "1.1 Indica todos los grupos legislativos y diputadas y diputados independientes (en caso de haberlos).      ",
                                    "cumple": "si",
                                    "errores": {},
                                    "comentario": ""
                                },
                                {
                                    "id": 2,
                                    "puntos": "1",
                                    "descripcion": "1.2  Lista completa de las y los diputados pertenecientes a cada grupo y fracción.",
                                    "cumple": "no",
                                    "errores": {"no_existe": True},
                                    "comentario": "No se localizó la lista de diputados"
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    ]
    return json.dumps(campos)

class ExportacionYValidacionTestCase(TestCase):
    def setUp(self):
        self.user = User.objects.create_user(username="testuser", password="password123")
        self.token = generate_access_token(self.user)
        self.client = Client()

        self.ev_capital = Evaluacion.objects.create(
            usuario=self.user,
            entidad="Municipio Test",
            periodo="2026-T1",
            tipo="Capital",
            datos_cuestionario=crear_datos_prueba_capital()
        )

        self.ev_congreso = Evaluacion.objects.create(
            usuario=self.user,
            entidad="Congreso Test",
            periodo="2026-LXI",
            tipo="Congreso",
            datos_cuestionario=crear_datos_prueba_congreso()
        )

    def test_validar_datos_cuestionario(self):
        res = validar_datos_cuestionario(self.ev_capital)
        self.assertTrue(res["valido"])
        self.assertEqual(res["stats"]["total_criterios"], 2)
        self.assertEqual(res["stats"]["criterios_si"], 1)
        self.assertEqual(res["stats"]["criterios_no"], 1)
        self.assertEqual(res["stats"]["deficiencias"]["no_existe"], 1)
        self.assertEqual(res["stats"]["total_observaciones"], 1)

    def test_generar_y_validar_excel_capital(self):
        wb, temp_path = generar_excel(self.ev_capital)
        self.assertTrue(os.path.exists(temp_path))

        ws = wb.active
        # Validar fila 2 (1.1 -> cumple=si -> Col J=1, Col F='X')
        self.assertEqual(ws.cell(2, 10).value, 1)
        self.assertEqual(ws.cell(2, 6).value, "X")
        # Validar fila 3 (1.2 -> cumple=no -> Col J=0, Col H='X', Col K=1 por no_existe)
        self.assertEqual(ws.cell(3, 10).value, 0)
        self.assertEqual(ws.cell(3, 8).value, "X")
        self.assertEqual(ws.cell(3, 11).value, 1)

        # Validar fórmulas en fila 28
        self.assertEqual(ws.cell(28, 4).value, "=SUM(D2:D27)")
        self.assertEqual(ws.cell(28, 10).value, "=SUM(J2:J27)/25")
        self.assertEqual(ws.cell(28, 11).value, "=SUM(K2:K27)/25")

        # Validar fórmulas en fila 187 (Calificación final)
        self.assertTrue("=D28+D48" in ws.cell(187, 4).value)
        self.assertTrue("=(J28+J48" in ws.cell(187, 10).value)

        # Validar filas de desglose 190-198
        self.assertEqual(ws.cell(190, 10).value, "=J28")
        self.assertEqual(ws.cell(191, 10).value, "=J48")

        # Ejecutar validador automático
        res = validar_exportacion_excel(self.ev_capital, wb)
        self.assertTrue(res["valido"], f"Errores encontrados: {res.get('errores')}")
        self.assertEqual(len(res["errores"]), 0)

        if os.path.exists(temp_path):
            os.remove(temp_path)

    def test_generar_y_validar_excel_congreso(self):
        wb, temp_path = generar_excel(self.ev_congreso)
        self.assertTrue(os.path.exists(temp_path))

        ws = wb.active
        # Encabezado en G1 debe contener la entidad
        self.assertIn("Congreso Test", ws.cell(1, 7).value)

        # Criterio 1.1 en fila 2 -> 1
        self.assertEqual(ws.cell(2, 7).value, 1)
        # Criterio 1.2 en fila 3 -> 0
        self.assertEqual(ws.cell(3, 7).value, 0)

        # Fórmula de bloque en fila 23
        self.assertEqual(ws.cell(23, 7).value, "=((SUM(G2:G22))/21)")
        # Fórmula de calificación final en fila 248
        self.assertTrue("=(G23+G84" in ws.cell(248, 7).value)
        # Desglose en filas 251-258
        self.assertEqual(ws.cell(251, 7).value, "=G23")

        # Ejecutar validador
        res = validar_exportacion_excel(self.ev_congreso, wb)
        self.assertTrue(res["valido"], f"Errores encontrados: {res.get('errores')}")
        self.assertEqual(len(res["errores"]), 0)

        if os.path.exists(temp_path):
            os.remove(temp_path)

    def test_generar_y_validar_word_observaciones(self):
        doc, temp_path = generar_word_observaciones(self.ev_capital)
        self.assertTrue(os.path.exists(temp_path))

        # Debe tener tabla de metadatos + 1 tarjeta de observación
        self.assertEqual(len(doc.tables), 2)

        res = validar_exportacion_word(self.ev_capital, doc, tipo_doc='observaciones')
        self.assertTrue(res["valido"], f"Errores en word obs: {res.get('errores')}")
        self.assertEqual(res["criterios_o_tarjetas_verificadas"], 1)

        if os.path.exists(temp_path):
            os.remove(temp_path)

    def test_generar_y_validar_word_plantilla_congreso(self):
        doc, temp_path = generar_word_plantilla(self.ev_congreso)
        self.assertTrue(os.path.exists(temp_path))

        res = validar_exportacion_word(self.ev_congreso, doc, tipo_doc='formato')
        self.assertTrue(res["valido"], f"Errores en word formato: {res.get('errores')}")

        if os.path.exists(temp_path):
            os.remove(temp_path)

    def test_validar_evaluacion_completa(self):
        res_cap = validar_evaluacion_completa(self.ev_capital)
        self.assertTrue(res_cap["valido"])
        self.assertEqual(len(res_cap["errores"]), 0)

        res_cong = validar_evaluacion_completa(self.ev_congreso)
        self.assertTrue(res_cong["valido"])
        self.assertEqual(len(res_cong["errores"]), 0)

    def test_api_validar_exportacion(self):
        response = self.client.get(
            f"/api/validar_exportacion?id={self.ev_capital.id}",
            HTTP_AUTHORIZATION=f"Bearer {self.token}"
        )
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["valido"])
        self.assertEqual(data["entidad"], "Municipio Test")

    def test_management_command_validar_exportaciones(self):
        out = StringIO()
        call_command('validar_exportaciones', stdout=out)
        output_str = out.getvalue()
        self.assertIn("Iniciando auditoría interna", output_str)
        self.assertIn("Municipio Test", output_str)
        self.assertIn("Congreso Test", output_str)
        self.assertIn("Todas las exportaciones validadas son 100% consistentes", output_str)
