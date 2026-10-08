import os
import re
import json
import tempfile
from datetime import datetime
import openpyxl
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from django.conf import settings

def normalize_text(text):
    if not text:
        return ""
    # Remover puntuaciones, espacios y convertir a minúsculas
    text = str(text).lower().strip()
    text = re.sub(r'[\s\n\r\t\.,;:_\-\(\)\[\]"\'\/]+', '', text)
    return text

def extract_criterios_map(datos_cuestionario):
    """Extrae un diccionario de criterios normalizados a partir del JSON de evaluación"""
    if isinstance(datos_cuestionario, str):
        try:
            campos = json.loads(datos_cuestionario)
        except Exception:
            campos = []
    else:
        campos = datos_cuestionario or []

    criterios_dict = {}
    total_criterios = 0
    criterios_si = 0
    criterios_no = 0
    criterios_na = 0
    deficiencias = {"no_existe": 0, "no_actualizada": 0, "no_corresponde": 0, "ilegible": 0, "enlace_roto": 0}
    observaciones_count = 0

    for campo in campos:
        for bloque in campo.get('bloques', []):
            for aspecto in bloque.get('aspectos', []):
                for c in aspecto.get('criterios', []):
                    total_criterios += 1
                    cumple = c.get('cumple')
                    if cumple == 'si':
                        criterios_si += 1
                    elif cumple == 'no':
                        criterios_no += 1
                    else:
                        criterios_na += 1

                    comentario = (c.get('comentario') or '').strip()
                    if comentario:
                        observaciones_count += 1

                    errs = c.get('errores', {}) or {}
                    for k in deficiencias.keys():
                        if errs.get(k):
                            deficiencias[k] += 1

                    desc = c.get('descripcion', '')
                    norm_key = normalize_text(desc)
                    if norm_key:
                        criterios_dict[norm_key] = c

    return {
        "criterios_dict": criterios_dict,
        "campos": campos,
        "stats": {
            "total_criterios": total_criterios,
            "criterios_si": criterios_si,
            "criterios_no": criterios_no,
            "criterios_na": criterios_na,
            "deficiencias": deficiencias,
            "total_observaciones": observaciones_count
        }
    }

def _set_cell_bg(cell, fill_hex):
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def generar_excel(evaluacion):
    """Genera el libro de Excel con fórmulas, datos y resultados estadísticos"""
    base_dir = getattr(settings, 'BASE_DIR', '/code')
    tipo = getattr(evaluacion, 'tipo', 'Capital')
    entidad = getattr(evaluacion, 'entidad', 'Entidad')
    periodo = getattr(evaluacion, 'periodo', 'Periodo')
    datos_cuestionario = getattr(evaluacion, 'datos_cuestionario', '[]')

    data_info = extract_criterios_map(datos_cuestionario)
    criterios_dict = data_info["criterios_dict"]

    plantilla_name = "Congreso.xlsx" if tipo == "Congreso" else "cimtra-oconca.xlsx"
    plantilla_path = os.path.join(base_dir, plantilla_name)
    if not os.path.exists(plantilla_path):
        raise FileNotFoundError(f"Plantilla {plantilla_name} no encontrada en {base_dir}")

    wb = openpyxl.load_workbook(plantilla_path, data_only=False)
    ws = wb.active

    if tipo == "Congreso":
        # Configurar encabezado con nombre de entidad en columna G
        ws.cell(1, 7).value = f"CIMTRA - {entidad} ({periodo})"

        for r in range(2, ws.max_row + 1):
            cell_desc = ws.cell(r, 5).value # Col E (Criterio)
            if cell_desc:
                norm_desc = normalize_text(cell_desc)
                matched_c = None
                if norm_desc in criterios_dict:
                    matched_c = criterios_dict[norm_desc]
                else:
                    # Intento de coincidencia parcial
                    for k, val in criterios_dict.items():
                        if k.startswith(norm_desc[:35]) or norm_desc.startswith(k[:35]):
                            matched_c = val
                            break

                if matched_c:
                    cumple = matched_c.get("cumple")
                    if cumple == "si":
                        ws.cell(r, 6).value = 1 # Col F
                        ws.cell(r, 7).value = 1 # Col G
                    elif cumple == "no":
                        ws.cell(r, 6).value = 0 # Col F
                        ws.cell(r, 7).value = 0 # Col G
                    else:
                        ws.cell(r, 6).value = 0
                        ws.cell(r, 7).value = 0

        # Asegurar y recalcular filas de fórmulas de Congreso
        # Bloques: 23 (2-22), 84 (24-83), 126 (85-125), 163 (127-162), 187 (164-186), 205 (188-204), 229 (206-228), 246 (230-245)
        bloques_congreso = [
            (23, 2, 22, 21),
            (84, 24, 83, 60),
            (126, 85, 125, 41),
            (163, 127, 162, 36),
            (187, 164, 186, 23),
            (205, 188, 204, 17),
            (229, 206, 228, 23),
            (246, 230, 245, 16)
        ]
        for row_idx, start_r, end_r, divisor in bloques_congreso:
            ws.cell(row_idx, 6).value = f"=((SUM(F{start_r}:F{end_r}))/{divisor})"
            ws.cell(row_idx, 7).value = f"=((SUM(G{start_r}:G{end_r}))/{divisor})"

        # Fila 248: Calificación final
        ws.cell(248, 4).value = "=D23+D84+D126+D163+D187+D205+D229+D246"
        ws.cell(248, 6).value = "=(F23+F84+F126+F163+F187+F205+F229+F246)/8"
        ws.cell(248, 7).value = "=(G23+G84+G126+G163+G187+G205+G229+G246)/8"

        # Filas 251-258: Desglose
        desglose_congreso = [(251, 23), (252, 84), (253, 126), (254, 163), (255, 187), (256, 205), (257, 229), (258, 246)]
        for r_target, r_src in desglose_congreso:
            ws.cell(r_target, 6).value = f"=F{r_src}"
            ws.cell(r_target, 7).value = f"=G{r_src}"

    else:
        # CIMTRA Capital / Municipal
        # Limpiar y poblar reactivos
        for r in range(2, 186):
            cell_desc = ws.cell(r, 5).value # Col E (Criterio)
            if cell_desc:
                norm_desc = normalize_text(cell_desc)
                matched_c = None
                if norm_desc in criterios_dict:
                    matched_c = criterios_dict[norm_desc]
                else:
                    for k, val in criterios_dict.items():
                        if k.startswith(norm_desc[:35]) or norm_desc.startswith(k[:35]):
                            matched_c = val
                            break

                # Limpiar columnas F, H, J, K, L, M, N, O
                ws.cell(r, 6).value = None  # Col F (Si)
                ws.cell(r, 8).value = None  # Col H (No)
                ws.cell(r, 10).value = 0    # Col J (Puntos Criterio)
                ws.cell(r, 11).value = 0    # Col K (No existe)
                ws.cell(r, 12).value = 0    # Col L (No actualizada)
                ws.cell(r, 13).value = 0    # Col M (No corresponde)
                ws.cell(r, 14).value = 0    # Col N (Ilegible)
                ws.cell(r, 15).value = 0    # Col O (Enlace roto)

                if matched_c:
                    cumple = matched_c.get("cumple")
                    if cumple == "si":
                        ws.cell(r, 6).value = "X"
                        ws.cell(r, 10).value = 1
                    elif cumple == "no":
                        ws.cell(r, 8).value = "X"
                        ws.cell(r, 10).value = 0

                    errs = matched_c.get("errores", {}) or {}
                    if errs.get("no_existe"): ws.cell(r, 11).value = 1
                    if errs.get("no_actualizada"): ws.cell(r, 12).value = 1
                    if errs.get("no_corresponde"): ws.cell(r, 13).value = 1
                    if errs.get("ilegible"): ws.cell(r, 14).value = 1
                    if errs.get("enlace_roto"): ws.cell(r, 15).value = 1

        # Fórmulas de bloques en Capital
        bloques_capital = [
            (28, 2, 27, 25),
            (48, 29, 47, 17),
            (73, 49, 72, 17),
            (120, 74, 119, 39),
            (127, 121, 126, 5),
            (135, 128, 134, 14),
            (148, 136, 147, 12),
            (167, 149, 166, 16),
            (185, 168, 184, 16)
        ]

        for row_idx, start_r, end_r, divisor in bloques_capital:
            ws.cell(row_idx, 4).value = f"=SUM(D{start_r}:D{end_r})"
            for col_idx, col_letter in [(10, 'J'), (11, 'K'), (12, 'L'), (13, 'M'), (14, 'N'), (15, 'O')]:
                ws.cell(row_idx, col_idx).value = f"=SUM({col_letter}{start_r}:{col_letter}{end_r})/{divisor}"

        # Fila 187: Calificación Final
        ws.cell(187, 4).value = "=D28+D48+D73+D120+D127+D135+D148+D167+D185"
        for col_idx, col_letter in [(10, 'J'), (11, 'K'), (12, 'L'), (13, 'M'), (14, 'N'), (15, 'O')]:
            ws.cell(187, col_idx).value = f"=({col_letter}28+{col_letter}48+{col_letter}73+{col_letter}120+{col_letter}127+{col_letter}135+{col_letter}148+{col_letter}167+{col_letter}185)/9"

        # Filas 190-198: Resultados por bloques
        desglose_capital = [
            (190, 28),
            (191, 48),
            (192, 73),
            (193, 120),
            (194, 127),
            (195, 135),
            (196, 148),
            (197, 167),
            (198, 185)
        ]
        for r_target, r_src in desglose_capital:
            for col_idx, col_letter in [(10, 'J'), (11, 'K'), (12, 'L'), (13, 'M'), (14, 'N'), (15, 'O')]:
                ws.cell(r_target, col_idx).value = f"={col_letter}{r_src}"

    temp_fd, temp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(temp_fd)
    wb.save(temp_path)
    return wb, temp_path


def generar_word_plantilla(evaluacion):
    """Genera el documento Word basado en la plantilla oficial CIMTRA-Legislativo"""
    base_dir = getattr(settings, 'BASE_DIR', '/code')
    entidad = getattr(evaluacion, 'entidad', 'Entidad')
    periodo = getattr(evaluacion, 'periodo', 'Periodo')
    usuario_nombre = evaluacion.usuario.username if getattr(evaluacion, 'usuario', None) else "Evaluador"
    datos_cuestionario = getattr(evaluacion, 'datos_cuestionario', '[]')

    plantilla_path = os.path.join(base_dir, "CIMTRA-Legislativo.docx")
    if not os.path.exists(plantilla_path):
        raise FileNotFoundError(f"Plantilla Word no encontrada en {plantilla_path}")

    data_info = extract_criterios_map(datos_cuestionario)
    criterios_dict = data_info["criterios_dict"]

    doc = docx.Document(plantilla_path)

    # Actualizar Tabla 0 (Metadatos)
    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        if len(t0.rows) >= 4:
            t0.rows[0].cells[0].text = f"Estado: {entidad}"
            t0.rows[0].cells[1].text = f"Fecha de aplicación: {datetime.now().strftime('%d/%m/%Y')}"
            t0.rows[1].cells[0].text = f"Periodo legislativo: {periodo}"
            t0.rows[1].cells[1].text = f"Aplicadores: {usuario_nombre}"

    for table in doc.tables[1:]:
        for row in table.rows:
            if len(row.cells) >= 5:
                desc = row.cells[0].text
                norm_desc = normalize_text(desc)
                matched_c = None
                if norm_desc in criterios_dict:
                    matched_c = criterios_dict[norm_desc]
                else:
                    for k, val in criterios_dict.items():
                        if k.startswith(norm_desc[:35]) or norm_desc.startswith(k[:35]):
                            matched_c = val
                            break

                if matched_c:
                    row.cells[2].text = ""
                    row.cells[4].text = ""
                    cumple = matched_c.get("cumple")
                    if cumple == "si":
                        row.cells[2].text = "X"
                    elif cumple == "no":
                        row.cells[4].text = "X"

    temp_fd, temp_path = tempfile.mkstemp(suffix=".docx")
    os.close(temp_fd)
    doc.save(temp_path)
    return doc, temp_path


def generar_word_observaciones(evaluacion):
    """Genera el informe formal de observaciones y no conformidades en Word"""
    entidad = getattr(evaluacion, 'entidad', 'Entidad')
    periodo = getattr(evaluacion, 'periodo', 'Periodo')
    tipo = getattr(evaluacion, 'tipo', 'Capital')
    usuario_nombre = evaluacion.usuario.username if getattr(evaluacion, 'usuario', None) else "Evaluador"
    datos_cuestionario = getattr(evaluacion, 'datos_cuestionario', '[]')

    try:
        campos = json.loads(datos_cuestionario) if isinstance(datos_cuestionario, str) else (datos_cuestionario or [])
    except Exception:
        campos = []

    doc = docx.Document()

    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    # Título Principal
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("OCONCA - SISTEMA DE EVALUACIÓN CIMTRA\n")
    run_title.font.name = 'Arial'
    run_title.font.size = Pt(16)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0, 31, 63)

    run_sub = p_title.add_run("INFORME DE OBSERVACIONES Y NO CONFORMIDADES")
    run_sub.font.name = 'Arial'
    run_sub.font.size = Pt(13)
    run_sub.font.bold = True
    run_sub.font.color.rgb = RGBColor(100, 100, 100)

    # Tabla de Datos de la Revisión
    table_meta = doc.add_table(rows=4, cols=2)
    table_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_meta.autofit = False

    meta_items = [
        ("Entidad Evaluada:", entidad),
        ("Período Evaluado:", periodo),
        ("Tipo de Evaluación:", f"CIMTRA {tipo}"),
        ("Evaluador / Fecha:", f"{usuario_nombre} | {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    ]

    for idx, (label, val) in enumerate(meta_items):
        row = table_meta.rows[idx]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(4.6)
        _set_cell_bg(cell_lbl, "F0F4F8")
        _set_cell_bg(cell_val, "FFFFFF")

        p_lbl = cell_lbl.paragraphs[0]
        r_lbl = p_lbl.add_run(label)
        r_lbl.font.name = 'Arial'
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(9.5)
        r_lbl.font.color.rgb = RGBColor(0, 31, 63)

        p_val = cell_val.paragraphs[0]
        r_val = p_val.add_run(val)
        r_val.font.name = 'Arial'
        r_val.font.size = Pt(9.5)

    doc.add_paragraph()

    criterios_observados = []
    total_observaciones = 0

    for campo in campos:
        campo_nom = campo.get('nombre', '')
        for bloque in campo.get('bloques', []):
            bloque_nom = bloque.get('nombre', '')
            for aspecto in bloque.get('aspectos', []):
                asp_titulo = aspecto.get('titulo', '')
                for c in aspecto.get('criterios', []):
                    cumple = c.get('cumple')
                    comentario = (c.get('comentario') or '').strip()
                    errores = c.get('errores', {}) or {}
                    tiene_errores = any([errores.get(k) for k in ['no_existe', 'no_actualizada', 'no_corresponde', 'ilegible', 'enlace_roto']])

                    if cumple == 'no' or comentario or tiene_errores:
                        total_observaciones += 1
                        criterios_observados.append({
                            'campo': campo_nom,
                            'bloque': bloque_nom,
                            'aspecto': asp_titulo,
                            'criterio': c,
                            'comentario': comentario,
                            'errores': errores,
                            'cumple': cumple
                        })

    p_res = doc.add_paragraph()
    r_res = p_res.add_run(f"Total de Puntos con Observación o Incumplimiento: {total_observaciones}")
    r_res.font.name = 'Arial'
    r_res.font.bold = True
    r_res.font.size = Pt(11)
    r_res.font.color.rgb = RGBColor(192, 57, 43)

    if total_observaciones == 0:
        p_empty = doc.add_paragraph()
        r_empty = p_empty.add_run("No se registraron observaciones ni puntos no cumplidos en esta evaluación.")
        r_empty.font.name = 'Arial'
        r_empty.font.italic = True
    else:
        ultimo_campo = None
        ultimo_bloque = None

        for item in criterios_observados:
            c = item['criterio']
            errs = item['errores']
            comentario = item['comentario']
            cumple = item['cumple']

            if item['campo'] != ultimo_campo:
                ultimo_campo = item['campo']
                p_c = doc.add_paragraph()
                p_c.paragraph_format.space_before = Pt(12)
                p_c.paragraph_format.space_after = Pt(4)
                r_c = p_c.add_run(f"📁 CAMPO: {str(ultimo_campo).upper()}")
                r_c.font.name = 'Arial'
                r_c.font.bold = True
                r_c.font.size = Pt(12)
                r_c.font.color.rgb = RGBColor(0, 31, 63)

            if item['bloque'] != ultimo_bloque:
                ultimo_bloque = item['bloque']
                p_b = doc.add_paragraph()
                p_b.paragraph_format.space_before = Pt(6)
                p_b.paragraph_format.space_after = Pt(4)
                r_b = p_b.add_run(f"  📂 {ultimo_bloque}")
                r_b.font.name = 'Arial'
                r_b.font.bold = True
                r_b.font.size = Pt(10.5)
                r_b.font.color.rgb = RGBColor(52, 73, 94)

            t_crit = doc.add_table(rows=3, cols=2)
            t_crit.alignment = WD_TABLE_ALIGNMENT.CENTER
            t_crit.autofit = False

            cell_top = t_crit.rows[0].cells[0]
            cell_top.merge(t_crit.rows[0].cells[1])
            _set_cell_bg(cell_top, "EAECEE")
            p_top = cell_top.paragraphs[0]
            
            r_asp = p_top.add_run(f"Aspecto: {item['aspecto']}\n")
            r_asp.font.name = 'Arial'
            r_asp.font.size = Pt(8.5)
            r_asp.font.italic = True
            r_asp.font.color.rgb = RGBColor(100, 100, 100)

            r_desc = p_top.add_run(f"Criterio: {c.get('descripcion', '')}")
            r_desc.font.name = 'Arial'
            r_desc.font.bold = True
            r_desc.font.size = Pt(9.5)
            r_desc.font.color.rgb = RGBColor(20, 20, 20)

            cell_cal = t_crit.rows[1].cells[0]
            cell_err = t_crit.rows[1].cells[1]
            cell_cal.width = Inches(2.2)
            cell_err.width = Inches(4.6)
            _set_cell_bg(cell_cal, "FAFAFA")
            _set_cell_bg(cell_err, "FAFAFA")

            p_cal = cell_cal.paragraphs[0]
            r_cal_lbl = p_cal.add_run("Calificación:\n")
            r_cal_lbl.font.name = 'Arial'
            r_cal_lbl.font.bold = True
            r_cal_lbl.font.size = Pt(8.5)
            
            estado_txt = "NO CUMPLE (0 pts)" if cumple == 'no' else ("CUMPLE" if cumple == 'si' else "N/A")
            r_cal_val = p_cal.add_run(estado_txt)
            r_cal_val.font.name = 'Arial'
            r_cal_val.font.bold = True
            r_cal_val.font.size = Pt(9.5)
            r_cal_val.font.color.rgb = RGBColor(192, 57, 43) if cumple == 'no' else RGBColor(39, 174, 96)

            p_err = cell_err.paragraphs[0]
            r_err_lbl = p_err.add_run("Deficiencias detectadas:\n")
            r_err_lbl.font.name = 'Arial'
            r_err_lbl.font.bold = True
            r_err_lbl.font.size = Pt(8.5)

            lista_err = []
            if errs.get('no_existe'): lista_err.append("No existe")
            if errs.get('no_actualizada'): lista_err.append("No actualizada")
            if errs.get('ilegible'): lista_err.append("Ilegible")
            if errs.get('enlace_roto'): lista_err.append("Enlace roto")
            if errs.get('no_corresponde'): lista_err.append("No corresponde")

            texto_errs = ", ".join(lista_err) if lista_err else "Ninguna deficiencia premarcada"
            r_err_val = p_err.add_run(texto_errs)
            r_err_val.font.name = 'Arial'
            r_err_val.font.size = Pt(9)
            r_err_val.font.italic = not bool(lista_err)

            cell_obs = t_crit.rows[2].cells[0]
            cell_obs.merge(t_crit.rows[2].cells[1])
            _set_cell_bg(cell_obs, "FCF3CF" if comentario else "FFFFFF")

            p_obs = cell_obs.paragraphs[0]
            r_obs_lbl = p_obs.add_run("Observación / Comentario:\n")
            r_obs_lbl.font.name = 'Arial'
            r_obs_lbl.font.bold = True
            r_obs_lbl.font.size = Pt(9)
            r_obs_lbl.font.color.rgb = RGBColor(125, 102, 8) if comentario else RGBColor(100, 100, 100)

            r_obs_val = p_obs.add_run(comentario if comentario else "(Sin observación específica redactada)")
            r_obs_val.font.name = 'Arial'
            r_obs_val.font.size = Pt(9.5)
            r_obs_val.font.italic = not bool(comentario)

            p_space = doc.add_paragraph()
            p_space.paragraph_format.space_before = Pt(0)
            p_space.paragraph_format.space_after = Pt(6)

    temp_fd, temp_path = tempfile.mkstemp(suffix=".docx")
    os.close(temp_fd)
    doc.save(temp_path)
    return doc, temp_path
