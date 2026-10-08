import os
import json
import openpyxl
import docx

from .exporter import (
    normalize_text,
    extract_criterios_map,
    generar_excel,
    generar_word_plantilla,
    generar_word_observaciones
)

def validar_datos_cuestionario(evaluacion):
    """Valida la consistencia del JSON de la evaluación"""
    datos = getattr(evaluacion, 'datos_cuestionario', '[]')
    info = extract_criterios_map(datos)
    stats = info["stats"]
    errores = []

    if stats["total_criterios"] == 0:
        errores.append("La evaluación no contiene criterios o reactivos válidos.")

    return {
        "valido": len(errores) == 0,
        "errores": errores,
        "stats": stats
    }


def validar_exportacion_excel(evaluacion, workbook_or_path=None):
    """
    Valida la coherencia de datos, fórmulas y resultados estadísticos en el archivo Excel generado.
    """
    temp_path = None
    if workbook_or_path is None:
        wb, temp_path = generar_excel(evaluacion)
    elif isinstance(workbook_or_path, str):
        wb = openpyxl.load_workbook(workbook_or_path, data_only=False)
    else:
        wb = workbook_or_path

    ws = wb.active
    tipo = getattr(evaluacion, 'tipo', 'Capital')
    datos_cuestionario = getattr(evaluacion, 'datos_cuestionario', '[]')
    data_info = extract_criterios_map(datos_cuestionario)
    criterios_dict = data_info["criterios_dict"]
    stats_internas = data_info["stats"]

    errores = []
    advertencias = []
    reactivos_verificados = 0
    puntos_sumados = 0

    if tipo == "Congreso":
        # Validar Congreso.xlsx
        header_g1 = str(ws.cell(1, 7).value or '')
        if getattr(evaluacion, 'entidad', '') and evaluacion.entidad not in header_g1:
            advertencias.append(f"El encabezado G1 '{header_g1}' no contiene el nombre de la entidad '{evaluacion.entidad}'")

        # Bloques de Congreso: (fila_formula, inicio_fila, fin_fila, divisor, nombre_bloque)
        bloques_cfg = [
            (23, 2, 22, 21, "Integración y estructura"),
            (84, 24, 83, 60, "Desempeño Legislativo"),
            (126, 85, 125, 41, "Documentación legislativa"),
            (163, 127, 162, 36, "Gastos"),
            (187, 164, 186, 23, "Administración"),
            (205, 188, 204, 17, "Control interno"),
            (229, 206, 228, 23, "Vinculación ciudadana"),
            (246, 230, 245, 16, "Acceso a información pública")
        ]

        resumen_bloques = []
        suma_promedios_bloques = 0.0

        for row_formula, r_ini, r_fin, divisor, nombre_b in bloques_cfg:
            pts_bloque = 0
            for r in range(r_ini, r_fin + 1):
                cell_desc = ws.cell(r, 5).value # Col E
                if cell_desc:
                    norm = normalize_text(cell_desc)
                    matched_c = None
                    if norm in criterios_dict:
                        matched_c = criterios_dict[norm]
                    else:
                        for k, val in criterios_dict.items():
                            if k.startswith(norm[:35]) or norm.startswith(k[:35]):
                                matched_c = val
                                break

                    if matched_c:
                        reactivos_verificados += 1
                        val_g = ws.cell(r, 7).value
                        cumple = matched_c.get("cumple")
                        expected_val = 1 if cumple == "si" else 0

                        if val_g != expected_val:
                            errores.append(f"Fila {r} ({cell_desc[:25]}...): Valor en Col G ({val_g}) no coincide con estado '{cumple}' (esperado {expected_val})")

                        if val_g == 1:
                            pts_bloque += 1
                            puntos_sumados += 1

            promedio_esperado = round((pts_bloque / divisor) * 100, 2)
            suma_promedios_bloques += (pts_bloque / divisor)

            # Validar fórmula de bloque en Col G
            formula_g = str(ws.cell(row_formula, 7).value or '')
            if f"SUM(G{r_ini}:G{r_fin})" not in formula_g.replace(" ", ""):
                errores.append(f"Fila {row_formula} ({nombre_b}): Fórmula incorrecta '{formula_g}' (esperado SUM(G{r_ini}:G{r_fin}))")

            resumen_bloques.append({
                "bloque": nombre_b,
                "fila_formula": row_formula,
                "puntos_obtenidos": pts_bloque,
                "divisor": divisor,
                "porcentaje": promedio_esperado,
                "formula": formula_g
            })

        # Validar fila 248 (Calificación Final)
        formula_final_g = str(ws.cell(248, 7).value or '')
        if "246)/8" not in formula_final_g.replace(" ", ""):
            errores.append(f"Fila 248: Fórmula de calificación final incorrecta '{formula_final_g}'")

        calificacion_final_esperada = round((suma_promedios_bloques / len(bloques_cfg)) * 100, 2)

        # Validar desglose en filas 251-258
        desglose_cfg = [(251, 23), (252, 84), (253, 126), (254, 163), (255, 187), (256, 205), (257, 229), (258, 246)]
        for r_target, r_src in desglose_cfg:
            f_val = str(ws.cell(r_target, 7).value or '')
            if f"=G{r_src}" not in f_val.replace(" ", ""):
                errores.append(f"Fila {r_target}: Desglose de bloque no referencia a =G{r_src} (obtenido '{f_val}')")

    else:
        # Validar Capital / cimtra-oconca.xlsx
        bloques_capital_cfg = [
            (28, 2, 27, 25, "Gastos"),
            (48, 29, 47, 17, "Obras"),
            (73, 49, 72, 17, "Bienes y sus Usos"),
            (120, 74, 119, 39, "Administración"),
            (127, 121, 126, 5, "Urbanidad"),
            (135, 128, 134, 14, "Consejos"),
            (148, 136, 147, 12, "Participación Ciudadana"),
            (167, 149, 166, 16, "Cabildo"),
            (185, 168, 184, 16, "Atención Ciudadana")
        ]

        resumen_bloques = []
        suma_promedios_bloques = 0.0

        for row_formula, r_ini, r_fin, divisor, nombre_b in bloques_capital_cfg:
            pts_bloque = 0
            defic_bloque = {"no_existe": 0, "no_actualizada": 0, "no_corresponde": 0, "ilegible": 0, "enlace_roto": 0}

            for r in range(r_ini, r_fin + 1):
                cell_desc = ws.cell(r, 5).value # Col E
                if cell_desc:
                    norm = normalize_text(cell_desc)
                    matched_c = None
                    if norm in criterios_dict:
                        matched_c = criterios_dict[norm]
                    else:
                        for k, val in criterios_dict.items():
                            if k.startswith(norm[:35]) or norm.startswith(k[:35]):
                                matched_c = val
                                break

                    if matched_c:
                        reactivos_verificados += 1
                        cumple = matched_c.get("cumple")
                        val_j = ws.cell(r, 10).value # Col J (Puntos)
                        expected_j = 1 if cumple == "si" else 0

                        if val_j != expected_j:
                            errores.append(f"Fila {r} ({cell_desc[:25]}...): Valor en Col J ({val_j}) no coincide con '{cumple}' (esperado {expected_j})")

                        if val_j == 1:
                            pts_bloque += 1
                            puntos_sumados += 1

                        # Marcas visuales F / H
                        val_f = ws.cell(r, 6).value
                        val_h = ws.cell(r, 8).value
                        if cumple == "si" and val_f != "X":
                            errores.append(f"Fila {r}: Falta marca 'X' en Columna F (Si)")
                        if cumple == "no" and val_h != "X":
                            errores.append(f"Fila {r}: Falta marca 'X' en Columna H (No)")

                        # Deficiencias en K..O
                        errs = matched_c.get("errores", {}) or {}
                        cols_err = [(11, 'no_existe'), (12, 'no_actualizada'), (13, 'no_corresponde'), (14, 'ilegible'), (15, 'enlace_roto')]
                        for c_idx, err_key in cols_err:
                            c_val = ws.cell(r, c_idx).value
                            exp_c = 1 if errs.get(err_key) else 0
                            if c_val != exp_c:
                                errores.append(f"Fila {r}: Columna {c_idx} ({err_key}) valor {c_val} != esperado {exp_c}")
                            if exp_c == 1:
                                defic_bloque[err_key] += 1

            promedio_bloque = round((pts_bloque / divisor) * 100, 2)
            suma_promedios_bloques += (pts_bloque / divisor)

            # Validar fórmulas en fila_formula
            f_d = str(ws.cell(row_formula, 4).value or '')
            f_j = str(ws.cell(row_formula, 10).value or '')
            if f"SUM(D{r_ini}:D{r_fin})" not in f_d.replace(" ", ""):
                errores.append(f"Fila {row_formula} ({nombre_b}): Fórmula en Col D incorrecta '{f_d}'")
            if f"SUM(J{r_ini}:J{r_fin})/{divisor}" not in f_j.replace(" ", ""):
                errores.append(f"Fila {row_formula} ({nombre_b}): Fórmula en Col J incorrecta '{f_j}'")

            resumen_bloques.append({
                "bloque": nombre_b,
                "fila_formula": row_formula,
                "puntos_obtenidos": pts_bloque,
                "divisor": divisor,
                "porcentaje": promedio_bloque,
                "deficiencias": defic_bloque,
                "formula_puntos": f_j
            })

        # Validar fila 187 (Calificación Final)
        f_final_j = str(ws.cell(187, 10).value or '')
        if "185)/9" not in f_final_j.replace(" ", ""):
            errores.append(f"Fila 187: Fórmula de calificación final incorrecta '{f_final_j}'")

        calificacion_final_esperada = round((suma_promedios_bloques / len(bloques_capital_cfg)) * 100, 2)

        # Validar desglose en filas 190-198
        desglose_cap = [(190, 28), (191, 48), (192, 73), (193, 120), (194, 127), (195, 135), (196, 148), (197, 167), (198, 185)]
        for r_target, r_src in desglose_cap:
            f_val = str(ws.cell(r_target, 10).value or '')
            if f"=J{r_src}" not in f_val.replace(" ", ""):
                errores.append(f"Fila {r_target}: Desglose de bloque en Col J no es =J{r_src} (obtenido '{f_val}')")

    if temp_path and os.path.exists(temp_path):
        try:
            os.remove(temp_path)
        except Exception:
            pass

    return {
        "valido": len(errores) == 0,
        "errores": errores,
        "advertencias": advertencias,
        "total_reactivos_verificados": reactivos_verificados,
        "puntos_totales_obtenidos": puntos_sumados,
        "calificacion_calculada": calificacion_final_esperada,
        "resumen_bloques": resumen_bloques
    }


def validar_exportacion_word(evaluacion, docx_or_path=None, tipo_doc='observaciones'):
    """
    Valida la estructura y consistencia de datos en los documentos Word exportados.
    """
    temp_path = None
    if docx_or_path is None:
        if tipo_doc == 'formato':
            doc, temp_path = generar_word_plantilla(evaluacion)
        else:
            doc, temp_path = generar_word_observaciones(evaluacion)
    elif isinstance(docx_or_path, str):
        doc = docx.Document(docx_or_path)
    else:
        doc = docx_or_path

    datos_cuestionario = getattr(evaluacion, 'datos_cuestionario', '[]')
    data_info = extract_criterios_map(datos_cuestionario)
    stats = data_info["stats"]
    criterios_dict = data_info["criterios_dict"]

    errores = []
    advertencias = []
    criterios_marcados = 0

    if tipo_doc == 'formato':
        # Plantilla oficial legislativa
        if len(doc.tables) < 2:
            errores.append("El documento no contiene las tablas de reactivos de la plantilla oficial.")
        else:
            # Validar Tabla 0
            t0 = doc.tables[0]
            t0_text = " ".join([c.text for row in t0.rows for c in row.cells])
            if getattr(evaluacion, 'entidad', '') and evaluacion.entidad not in t0_text:
                advertencias.append(f"La tabla de metadatos no contiene la entidad '{evaluacion.entidad}'")

            for table in doc.tables[1:]:
                for row in table.rows:
                    if len(row.cells) >= 5:
                        desc = row.cells[0].text
                        norm = normalize_text(desc)
                        matched = None
                        if norm in criterios_dict:
                            matched = criterios_dict[norm]
                        else:
                            for k, val in criterios_dict.items():
                                if k.startswith(norm[:35]) or norm.startswith(k[:35]):
                                    matched = val
                                    break

                        if matched:
                            cumple = matched.get("cumple")
                            m_si = row.cells[2].text.strip()
                            m_no = row.cells[4].text.strip()

                            if cumple == "si":
                                if m_si != "X":
                                    errores.append(f"Criterio '{desc[:25]}...': Falta marca 'X' en columna 'Sí'")
                                criterios_marcados += 1
                            elif cumple == "no":
                                if m_no != "X":
                                    errores.append(f"Criterio '{desc[:25]}...': Falta marca 'X' en columna 'No'")
                                criterios_marcados += 1

    else:
        # Informe de Observaciones
        if len(doc.tables) == 0:
            errores.append("El informe de observaciones no contiene tablas.")
        else:
            # Validar metadata en tabla 0
            t0 = doc.tables[0]
            t0_text = " ".join([c.text for row in t0.rows for c in row.cells])
            if getattr(evaluacion, 'entidad', '') and evaluacion.entidad not in t0_text:
                errores.append(f"El informe no indica la entidad '{evaluacion.entidad}' en el encabezado.")

            # Validar número de observaciones listadas
            # Cada tarjeta de observación es una tabla de 3 filas
            tablas_observaciones = len(doc.tables) - 1 # Menos la de metadata
            
            # Contar total esperado de observaciones/incumplimientos
            total_esperado = 0
            for campo in data_info["campos"]:
                for b in campo.get('bloques', []):
                    for asp in b.get('aspectos', []):
                        for c in asp.get('criterios', []):
                            comentario = (c.get('comentario') or '').strip()
                            errs = c.get('errores', {}) or {}
                            tiene_errs = any([errs.get(k) for k in ['no_existe', 'no_actualizada', 'no_corresponde', 'ilegible', 'enlace_roto']])
                            if c.get('cumple') == 'no' or comentario or tiene_errs:
                                total_esperado += 1

            if tablas_observaciones != total_esperado:
                errores.append(f"Total de tarjetas de observaciones en Word ({tablas_observaciones}) no coincide con el total esperado ({total_esperado})")

            criterios_marcados = tablas_observaciones

    if temp_path and os.path.exists(temp_path):
        try:
            os.remove(temp_path)
        except Exception:
            pass

    return {
        "valido": len(errores) == 0,
        "errores": errores,
        "advertencias": advertencias,
        "total_tablas_documento": len(doc.tables),
        "criterios_o_tarjetas_verificadas": criterios_marcados
    }


def validar_evaluacion_completa(evaluacion):
    """
    Ejecuta el proceso integral de validación interna para una evaluación:
    - Validación de consistencia del JSON de evaluación
    - Validación de exportación a Excel con fórmulas y estadísticas
    - Validación de exportación a Word de Observaciones
    - Validación de exportación a Word de Plantilla oficial (si aplica Congreso)
    """
    res_json = validar_datos_cuestionario(evaluacion)
    res_excel = validar_exportacion_excel(evaluacion)
    res_word_obs = validar_exportacion_word(evaluacion, tipo_doc='observaciones')
    
    res_word_formato = None
    if getattr(evaluacion, 'tipo', 'Capital') == "Congreso":
        res_word_formato = validar_exportacion_word(evaluacion, tipo_doc='formato')

    es_valido = res_json["valido"] and res_excel["valido"] and res_word_obs["valido"]
    if res_word_formato:
        es_valido = es_valido and res_word_formato["valido"]

    todos_errores = []
    todos_errores.extend([f"[JSON] {e}" for e in res_json.get("errores", [])])
    todos_errores.extend([f"[EXCEL] {e}" for e in res_excel.get("errores", [])])
    todos_errores.extend([f"[WORD_OBS] {e}" for e in res_word_obs.get("errores", [])])
    if res_word_formato:
        todos_errores.extend([f"[WORD_FORMATO] {e}" for e in res_word_formato.get("errores", [])])

    todas_advertencias = []
    todas_advertencias.extend([f"[EXCEL] {a}" for a in res_excel.get("advertencias", [])])
    todas_advertencias.extend([f"[WORD_OBS] {a}" for a in res_word_obs.get("advertencias", [])])
    if res_word_formato:
        todas_advertencias.extend([f"[WORD_FORMATO] {a}" for a in res_word_formato.get("advertencias", [])])

    return {
        "valido": es_valido,
        "entidad": getattr(evaluacion, 'entidad', ''),
        "periodo": getattr(evaluacion, 'periodo', ''),
        "tipo": getattr(evaluacion, 'tipo', 'Capital'),
        "errores": todos_errores,
        "advertencias": todas_advertencias,
        "resumen_excel": {
            "reactivos_verificados": res_excel["total_reactivos_verificados"],
            "puntos_obtenidos": res_excel["puntos_totales_obtenidos"],
            "calificacion_final": res_excel["calificacion_calculada"],
            "bloques": res_excel["resumen_bloques"]
        },
        "resumen_word": {
            "observaciones_verificadas": res_word_obs["criterios_o_tarjetas_verificadas"],
            "formato_oficial_verificado": bool(res_word_formato and res_word_formato["valido"])
        }
    }
