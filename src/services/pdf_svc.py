"""
Servicio de generación de informes PDF en memoria usando `fpdf2` y `BytesIO`.
Adaptado para incluir toda la información de Resoluciones Docentes (UNAD):
  a) Ficha individual de un docente (`generar_pdf_ficha_docente`): Datos personales, resolución de vinculación,
     programa/centro/escuela, nivel de inglés, estudios y aportes con enlaces a Drive.
  b) Informe consolidado gerencial (`generar_pdf_informe_consolidado`): KPIs de contratación,
     distribución por nivel de inglés, modalidad, centro, formación y aportes.
"""

from datetime import datetime
from io import BytesIO
from typing import Any, Dict, List
from fpdf import FPDF


def _texto_latin1(texto: Any) -> str:
    """Normaliza cadenas Unicode a Latin-1 de manera segura para las fuentes core (Helvetica) de fpdf2."""
    if texto is None:
        return "-"
    cadena = str(texto)
    reemplazos = {
        "\u2013": "-",
        "\u2014": "-",
        "\u2018": "'",
        "\u2019": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u2022": "*",
        "\u2026": "...",
    }
    for orig, dest in reemplazos.items():
        cadena = cadena.replace(orig, dest)
    return cadena.encode("latin-1", errors="replace").decode("latin-1")


def _formatear_moneda(valor: Any) -> str:
    if valor is None:
        return "-"
    try:
        num = float(valor)
        return f"$ {num:,.0f}".replace(",", ".")
    except Exception:
        return str(valor)


class ReporteAcademicoPDF(FPDF):
    """Clase base FPDF con encabezado y pie de página institucional."""

    def __init__(self, titulo_reporte: str, subtitulo_reporte: str = "") -> None:
        super().__init__(orientation="P", unit="mm", format="A4")
        self.titulo_reporte = _texto_latin1(titulo_reporte)
        self.subtitulo_reporte = _texto_latin1(subtitulo_reporte)
        self.fecha_generacion = datetime.now().strftime("%Y-%m-%d %H:%M")
        self.set_auto_page_break(auto=True, margin=18)

    def header(self) -> None:
        self.set_fill_color(0, 75, 135)
        self.set_text_color(255, 255, 255)
        self.set_font("Helvetica", "B", 13)
        self.cell(
            0,
            10,
            _texto_latin1("SISTEMA DE GESTION ACADEMICA Y RESOLUCIONES DOCENTES"),
            border=0,
            new_x="LMARGIN",
            new_y="NEXT",
            align="C",
            fill=True,
        )

        self.set_text_color(30, 41, 59)
        self.set_font("Helvetica", "B", 11)
        self.ln(2)
        self.cell(
            0,
            6,
            self.titulo_reporte,
            border=0,
            new_x="LMARGIN",
            new_y="NEXT",
            align="C",
        )

        if self.subtitulo_reporte:
            self.set_font("Helvetica", "I", 9)
            self.set_text_color(100, 116, 139)
            self.cell(
                0,
                5,
                self.subtitulo_reporte,
                border=0,
                new_x="LMARGIN",
                new_y="NEXT",
                align="C",
            )

        self.set_font("Helvetica", "", 8)
        self.set_text_color(100, 116, 139)
        self.cell(
            0,
            5,
            _texto_latin1(f"Fecha de generacion: {self.fecha_generacion}"),
            border=0,
            new_x="LMARGIN",
            new_y="NEXT",
            align="R",
        )
        self.set_draw_color(203, 213, 225)
        self.line(10, self.get_y() + 1, 200, self.get_y() + 1)
        self.ln(4)

    def footer(self) -> None:
        self.set_y(-14)
        self.set_draw_color(203, 213, 225)
        self.line(10, self.get_y(), 200, self.get_y())
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(100, 116, 139)
        self.cell(
            0,
            8,
            _texto_latin1(f"Sistema de Gestion Academica | Pagina {self.page_no()}/{{nb}}"),
            align="C",
        )

    def seccion_titulo(self, texto: str) -> None:
        self.set_fill_color(241, 245, 249)
        self.set_text_color(0, 75, 135)
        self.set_font("Helvetica", "B", 10)
        self.cell(
            0,
            7,
            _texto_latin1(f"  {texto}"),
            border=1,
            new_x="LMARGIN",
            new_y="NEXT",
            align="L",
            fill=True,
        )
        self.ln(2)


def generar_pdf_ficha_docente(
    docente: Dict[str, Any],
    estudios: List[Dict[str, Any]],
    aportes: List[Dict[str, Any]],
) -> bytes:
    """
    Genera en memoria (`bytes`) la Ficha Individual / Hoja de Vida y Resolución de Vinculación de un docente.
    """
    nombre = docente.get("nombre_completo", "Docente")
    documento = docente.get("documento", "N/A")
    res_num = docente.get("resolucion_numero") or "S/N"

    pdf = ReporteAcademicoPDF(
        titulo_reporte=f"FICHA INDIVIDUAL DOCENTE: {nombre}",
        subtitulo_reporte=f"Documento: {documento} | Resolucion Nro: {res_num} | Periodo: {docente.get('periodo_academico') or '-'}",
    )
    pdf.alias_nb_pages()
    pdf.add_page()

    # 1. Datos Personales y Competencia en Inglés
    pdf.seccion_titulo("1. Identificacion, Contacto y Competencia en Ingles")
    pdf.set_text_color(30, 41, 59)

    detalle_ing = docente.get("detalle_ingles_original")
    nivel_ing_mostrar = docente.get("nivel_ingles", "Sin certificar")
    if detalle_ing and detalle_ing != nivel_ing_mostrar:
        nivel_ing_mostrar = f"{nivel_ing_mostrar} ({detalle_ing})"

    estado_cod = str(docente.get("estado") or "A").strip().upper()
    estado_txt = "Retirado (R)" if estado_cod.startswith("R") else "Activo (A)"

    campos_personales = [
        ("Nombre Completo:", docente.get("nombre_completo", "-")),
        ("Documento / Estado:", f"{docente.get('documento', '-')}  |  Estado: {estado_txt}"),
        ("Correo Institucional:", docente.get("correo_institucional", "-")),
        ("Correo Personal / Telefono:", f"{docente.get('email_personal') or '-'}  |  Tel: {docente.get('telefono') or '-'}"),
        ("Nivel Maximo de Formacion:", docente.get("nivel_formacion_maximo") or "-"),
        ("Nivel de Ingles (MCER):", nivel_ing_mostrar),
        (
            "Soporte Certificado Ingles:",
            "Registrado en Google Drive" if docente.get("soporte_ingles_url") else "Sin soporte adjunto",
        ),
    ]

    for etiqueta, valor in campos_personales:
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(52, 5.5, _texto_latin1(etiqueta), border=0)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.cell(0, 5.5, _texto_latin1(valor), border=0, new_x="LMARGIN", new_y="NEXT")

    if docente.get("soporte_ingles_url"):
        pdf.set_font("Helvetica", "I", 8)
        pdf.set_text_color(0, 75, 135)
        pdf.cell(
            0,
            5,
            _texto_latin1(f"Enlace Drive Ingles: {docente['soporte_ingles_url']}"),
            border=0,
            new_x="LMARGIN",
            new_y="NEXT",
            link=str(docente["soporte_ingles_url"]),
        )
        pdf.set_text_color(30, 41, 59)

    pdf.ln(3)

    # 2. Datos de Resolución y Vinculación Académica
    pdf.seccion_titulo("2. Datos de Resolucion y Vinculacion Academica")
    campos_vinculacion = [
        (
            "Resolucion / Fecha / Estado:",
            f"Res. {docente.get('resolucion_numero') or '-'} del {docente.get('resolucion_fecha') or '-'} "
            f"(Estado: {docente.get('estado_aceptacion') or '-'})",
        ),
        (
            "Periodo / Modalidad / Horas:",
            f"Periodo {docente.get('periodo_academico') or '-'} | {docente.get('modalidad') or '-'} | "
            f"{docente.get('horas_vinculacion') or 0} horas",
        ),
        (
            "Vigencia Vinculacion:",
            f"Inicio: {docente.get('fecha_inicio') or '-'}  ->  Termina: {docente.get('fecha_termina') or '-'} "
            f"(Ultima vinc.: {docente.get('ultima_vinculacion') or '-'})",
        ),
        ("Escuela:", f"{docente.get('escuela_descripcion') or '-'} (Cod. {docente.get('escuela_codigo') or '-'})"),
        ("Centro / Zona:", f"{docente.get('centro_descripcion') or '-'} | {docente.get('zona_descripcion') or '-'}"),
        ("Programa Academico:", f"{docente.get('programa_descripcion') or '-'} (Cod. {docente.get('programa_codigo') or '-'})"),
        (
            "Salario Mensual / Total:",
            f"Mensual: {_formatear_moneda(docente.get('salario_mensual'))}  |  "
            f"Total Vinculacion: {_formatear_moneda(docente.get('total_vinculacion'))}",
        ),
        ("DOFE / Docente Nuevo:", f"DOFE: {docente.get('dofe') or '-'}  |  Nuevo: {docente.get('es_nuevo') or '-'}"),
    ]

    for etiqueta, valor in campos_vinculacion:
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(52, 5.5, _texto_latin1(etiqueta), border=0)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.cell(0, 5.5, _texto_latin1(valor)[:95], border=0, new_x="LMARGIN", new_y="NEXT")

    pdf.ln(3)

    # 3. Formación Académica (Estudios)
    pdf.seccion_titulo(f"3. Formacion Academica ({len(estudios)} titulos registrados)")
    if not estudios:
        pdf.set_font("Helvetica", "I", 8.5)
        pdf.cell(
            0,
            6,
            _texto_latin1("No se han registrado estudios detallados para este docente."),
            new_x="LMARGIN",
            new_y="NEXT",
        )
    else:
        pdf.set_fill_color(0, 75, 135)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(28, 6.5, _texto_latin1("Nivel"), border=1, fill=True)
        pdf.cell(72, 6.5, _texto_latin1("Titulo Obtenido"), border=1, fill=True)
        pdf.cell(52, 6.5, _texto_latin1("Universidad"), border=1, fill=True)
        pdf.cell(14, 6.5, _texto_latin1("Anio"), border=1, align="C", fill=True)
        pdf.cell(24, 6.5, _texto_latin1("Soporte"), border=1, align="C", new_x="LMARGIN", new_y="NEXT", fill=True)

        pdf.set_text_color(30, 41, 59)
        pdf.set_font("Helvetica", "", 7.5)
        for est in estudios:
            nivel = _texto_latin1(est.get("nivel_formacion", "-"))[:16]
            titulo = _texto_latin1(est.get("titulo_obtenido", "-"))[:44]
            univ = _texto_latin1(est.get("universidad", "-"))[:31]
            anio = _texto_latin1(est.get("anio_graduacion") or "-")
            url_sop = est.get("soporte_url")

            pdf.cell(28, 6.5, nivel, border=1)
            pdf.cell(72, 6.5, titulo, border=1)
            pdf.cell(52, 6.5, univ, border=1)
            pdf.cell(14, 6.5, anio, border=1, align="C")
            if url_sop:
                pdf.set_text_color(0, 75, 135)
                pdf.cell(
                    24,
                    6.5,
                    _texto_latin1("Abrir Drive"),
                    border=1,
                    align="C",
                    new_x="LMARGIN",
                    new_y="NEXT",
                    link=str(url_sop),
                )
                pdf.set_text_color(30, 41, 59)
            else:
                pdf.cell(24, 6.5, _texto_latin1("No adjunto"), border=1, align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(4)

    # 4. Aportes Académicos
    pdf.seccion_titulo(f"4. Aportes Academicos e Institucionales ({len(aportes)} registrados)")
    if not aportes:
        pdf.set_font("Helvetica", "I", 8.5)
        pdf.cell(
            0,
            6,
            _texto_latin1("No se han registrado aportes para este docente."),
            new_x="LMARGIN",
            new_y="NEXT",
        )
    else:
        for idx, ap in enumerate(aportes, start=1):
            cat = _texto_latin1(ap.get("categoria", "-"))
            tit = _texto_latin1(ap.get("titulo_aporte", "-"))
            per = _texto_latin1(ap.get("periodo_academico", "-"))
            desc = _texto_latin1(ap.get("descripcion") or "Sin descripcion adicional.")
            url_ap = ap.get("soporte_url")

            pdf.set_font("Helvetica", "B", 8.5)
            pdf.cell(
                0,
                5.5,
                _texto_latin1(f"{idx}. [{cat} | Periodo {per}] {tit}"),
                new_x="LMARGIN",
                new_y="NEXT",
            )
            pdf.set_font("Helvetica", "", 8)
            pdf.multi_cell(0, 5, _texto_latin1(f"   Descripcion: {desc}"), new_x="LMARGIN", new_y="NEXT")
            if url_ap:
                pdf.set_font("Helvetica", "I", 8)
                pdf.set_text_color(0, 75, 135)
                pdf.cell(
                    0,
                    5,
                    _texto_latin1(f"   Soporte Drive: {url_ap}"),
                    new_x="LMARGIN",
                    new_y="NEXT",
                    link=str(url_ap),
                )
                pdf.set_text_color(30, 41, 59)
            pdf.ln(1.5)

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer.getvalue()


def generar_pdf_informe_consolidado(
    stats: Dict[str, Any],
    docentes_resumen: List[Dict[str, Any]],
) -> bytes:
    """
    Genera en memoria (`bytes`) el Informe Consolidado Gerencial con estadísticas de resoluciones,
    niveles de inglés, centros, formación y aportes.
    """
    pdf = ReporteAcademicoPDF(
        titulo_reporte="INFORME CONSOLIDADO GERENCIAL - RESOLUCIONES Y GESTION DOCENTE",
        subtitulo_reporte="Resumen Estadistico de Vinculacion, Ingles (MCER), Formacion de Posgrado y Aportes",
    )
    pdf.alias_nb_pages()
    pdf.add_page()

    total_docentes = max(int(stats.get("total_docentes", 0)), 0)
    total_retirados = int(stats.get("total_docentes_retirados", 0))
    total_estudios = int(stats.get("total_estudios", 0))
    total_posgrados = int(stats.get("total_posgrados", 0))
    total_aportes = int(stats.get("total_aportes", 0))
    con_soporte_ing = int(stats.get("docentes_con_soporte_ingles", 0))
    total_horas = int(stats.get("total_horas_vinculacion", 0))
    suma_vinculacion = float(stats.get("suma_total_vinculacion", 0.0))

    # 1. Indicadores Clave
    pdf.seccion_titulo("1. Indicadores Clave del Cuerpo Docente Activo (Excluye Retirados 'R')")
    pdf.set_text_color(30, 41, 59)

    kpis = [
        ("Total de Docentes Activos ('A'):", f"{total_docentes} (Excluidos en estado Retirado 'R': {total_retirados})"),
        ("Total Horas Contratadas en Resoluciones:", f"{total_horas:,}".replace(",", ".")),
        ("Valor Total Consolidado de Vinculacion:", _formatear_moneda(suma_vinculacion)),
        ("Docentes con Certificado de Ingles en Drive:", f"{con_soporte_ing} de {total_docentes}"),
        ("Total de Titulos Academicos Desglosados:", f"{total_estudios} ({total_posgrados} de Posgrado)"),
        ("Total de Aportes Academicos Registrados:", str(total_aportes)),
    ]
    for k, v in kpis:
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(90, 5.5, _texto_latin1(k))
        pdf.set_font("Helvetica", "", 8.5)
        pdf.cell(0, 5.5, _texto_latin1(v), new_x="LMARGIN", new_y="NEXT")

    pdf.ln(3)

    # 2. Tabla de Distribución por Nivel de Inglés
    pdf.seccion_titulo("2. Distribucion de Docentes por Nivel de Ingles (MCER)")
    pdf.set_fill_color(0, 75, 135)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(90, 6.5, _texto_latin1("Nivel de Ingles"), border=1, fill=True)
    pdf.cell(50, 6.5, _texto_latin1("Cantidad de Docentes"), border=1, align="C", fill=True)
    pdf.cell(50, 6.5, _texto_latin1("Participacion (%)"), border=1, align="C", new_x="LMARGIN", new_y="NEXT", fill=True)

    pdf.set_text_color(30, 41, 59)
    pdf.set_font("Helvetica", "", 8.5)
    dist_ingles: Dict[str, int] = stats.get("distribucion_ingles", {})
    for nivel, cant in dist_ingles.items():
        porcentaje = (cant / total_docentes * 100.0) if total_docentes > 0 else 0.0
        pdf.cell(90, 6, _texto_latin1(nivel), border=1)
        pdf.cell(50, 6, str(cant), border=1, align="C")
        pdf.cell(50, 6, f"{porcentaje:.1f}%", border=1, align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(3)

    # 3. Tabla de Distribución por Centro y Modalidad
    dist_centro: Dict[str, int] = stats.get("distribucion_centro", {})
    if dist_centro:
        pdf.seccion_titulo("3. Distribucion de Docentes por Centro (CEAD / CCAV / UDR)")
        pdf.set_fill_color(0, 75, 135)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(110, 6.5, _texto_latin1("Centro"), border=1, fill=True)
        pdf.cell(40, 6.5, _texto_latin1("Docentes"), border=1, align="C", fill=True)
        pdf.cell(40, 6.5, _texto_latin1("Participacion (%)"), border=1, align="C", new_x="LMARGIN", new_y="NEXT", fill=True)

        pdf.set_text_color(30, 41, 59)
        pdf.set_font("Helvetica", "", 8.5)
        for cen, cant in sorted(dist_centro.items(), key=lambda x: x[1], reverse=True):
            porcentaje = (cant / total_docentes * 100.0) if total_docentes > 0 else 0.0
            pdf.cell(110, 6, _texto_latin1(cen)[:60], border=1)
            pdf.cell(40, 6, str(cant), border=1, align="C")
            pdf.cell(40, 6, f"{porcentaje:.1f}%", border=1, align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)

    # 4. Tabla de Aportes por Categoría
    pdf.seccion_titulo("4. Consolidado de Aportes por Categoria Institucional")
    pdf.set_fill_color(0, 75, 135)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(90, 6.5, _texto_latin1("Categoria de Aporte"), border=1, fill=True)
    pdf.cell(50, 6.5, _texto_latin1("Total Aportes"), border=1, align="C", fill=True)
    pdf.cell(50, 6.5, _texto_latin1("Participacion (%)"), border=1, align="C", new_x="LMARGIN", new_y="NEXT", fill=True)

    pdf.set_text_color(30, 41, 59)
    pdf.set_font("Helvetica", "", 8.5)
    dist_aportes: Dict[str, int] = stats.get("distribucion_aportes", {})
    for cat, cant in dist_aportes.items():
        porcentaje = (cant / total_aportes * 100.0) if total_aportes > 0 else 0.0
        pdf.cell(90, 6, _texto_latin1(cat), border=1)
        pdf.cell(50, 6, str(cant), border=1, align="C")
        pdf.cell(50, 6, f"{porcentaje:.1f}%", border=1, align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(4)

    # 5. Listado Resumido de Docentes en la Consulta
    if docentes_resumen:
        pdf.seccion_titulo(f"5. Listado de Docentes en la Consulta Actual ({len(docentes_resumen)} registros)")
        pdf.set_fill_color(0, 75, 135)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.cell(23, 6.5, _texto_latin1("Documento"), border=1, fill=True)
        pdf.cell(62, 6.5, _texto_latin1("Nombre Completo"), border=1, fill=True)
        pdf.cell(24, 6.5, _texto_latin1("Formacion"), border=1, align="C", fill=True)
        pdf.cell(22, 6.5, _texto_latin1("Ingles"), border=1, align="C", fill=True)
        pdf.cell(39, 6.5, _texto_latin1("Centro"), border=1, fill=True)
        pdf.cell(20, 6.5, _texto_latin1("Res. / Hrs"), border=1, align="C", new_x="LMARGIN", new_y="NEXT", fill=True)

        pdf.set_text_color(30, 41, 59)
        pdf.set_font("Helvetica", "", 7)
        for d in docentes_resumen[:120]:
            doc_num = _texto_latin1(d.get("documento", "-"))[:14]
            nom = _texto_latin1(d.get("nombre_completo", "-"))[:36]
            form = _texto_latin1(d.get("nivel_formacion_maximo") or "-")[:14]
            niv = _texto_latin1(d.get("nivel_ingles", "-"))[:12]
            cen = _texto_latin1(d.get("centro_descripcion") or "-")[:22]
            res_h = _texto_latin1(f"{d.get('resolucion_numero') or '-'} ({d.get('horas_vinculacion') or 0}h)")[:13]

            pdf.cell(23, 6, doc_num, border=1)
            pdf.cell(62, 6, nom, border=1)
            pdf.cell(24, 6, form, border=1, align="C")
            pdf.cell(22, 6, niv, border=1, align="C")
            pdf.cell(39, 6, cen, border=1)
            pdf.cell(20, 6, res_h, border=1, align="C", new_x="LMARGIN", new_y="NEXT")

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer.getvalue()


def generar_pdf_informe_matriculas(
    stats: Dict[str, Any],
    filtro_anio: str = "Todos",
) -> bytes:
    """
    Genera en memoria (`bytes`) el Informe Ejecutivo Consolidado de Matrículas (Actas 2025 y 2026)
    con desglose por Año, Periodo, Centro (CEAD/CCAV/UDR), Programa Académico y Condición (Nuevos vs Antiguos).
    """
    subtitulo = (
        f"Consolidado General de Actas de Matricula (Filtro Anio: {filtro_anio})"
        if filtro_anio != "Todos"
        else "Comparativo Consolidado de Actas de Matricula 2025 y 2026 (ECISA - ZCBC)"
    )
    pdf = ReporteAcademicoPDF(
        titulo_reporte="INFORME EJECUTIVO DE MATRICULAS POR PERIODO, CENTRO Y PROGRAMA",
        subtitulo_reporte=subtitulo,
    )
    pdf.alias_nb_pages()
    pdf.add_page()

    total_mat = max(int(stats.get("total_matriculas", 0)), 0)
    est_unicos = int(stats.get("estudiantes_unicos", 0))
    nuevos = int(stats.get("total_nuevos", 0))
    antiguos = int(stats.get("total_antiguos", 0))
    otros = int(stats.get("total_continuidad_otros", 0))

    # 1. KPIs Globales
    pdf.seccion_titulo("1. Resumen Ejecutivo de Matricula")
    pdf.set_text_color(30, 41, 59)
    kpis = [
        ("Total Actas de Matricula Registradas:", f"{total_mat:,}".replace(",", ".")),
        ("Total Estudiantes Unicos (por Documento):", f"{est_unicos:,}".replace(",", ".")),
        (
            "Estudiantes Antiguos:",
            f"{antiguos:,} ({(antiguos / total_mat * 100):.1f}%)".replace(",", ".") if total_mat else "0",
        ),
        (
            "Estudiantes Nuevos:",
            f"{nuevos:,} ({(nuevos / total_mat * 100):.1f}%)".replace(",", ".") if total_mat else "0",
        ),
        ("Continuidad / Aplazados / Otros:", f"{otros:,}".replace(",", ".")),
    ]
    for k, v in kpis:
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(85, 5.5, _texto_latin1(k))
        pdf.set_font("Helvetica", "", 8.5)
        pdf.cell(0, 5.5, _texto_latin1(v), new_x="LMARGIN", new_y="NEXT")

    pdf.ln(3)

    # 2. Comparativo por Año y Condición
    por_anio: Dict[str, int] = stats.get("por_anio", {})
    if por_anio:
        pdf.seccion_titulo("2. Evolucion de Matriculas por Anio de Acta")
        pdf.set_fill_color(0, 75, 135)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(70, 6.5, _texto_latin1("Anio del Acta"), border=1, fill=True)
        pdf.cell(60, 6.5, _texto_latin1("Total Matriculas"), border=1, align="C", fill=True)
        pdf.cell(60, 6.5, _texto_latin1("Participacion (%)"), border=1, align="C", new_x="LMARGIN", new_y="NEXT", fill=True)

        pdf.set_text_color(30, 41, 59)
        pdf.set_font("Helvetica", "", 8.5)
        for anio, cant in sorted(por_anio.items()):
            pct = (cant / total_mat * 100.0) if total_mat else 0.0
            pdf.cell(70, 6, _texto_latin1(f"Actas {anio}"), border=1)
            pdf.cell(60, 6, f"{cant:,}".replace(",", "."), border=1, align="C")
            pdf.cell(60, 6, f"{pct:.1f}%", border=1, align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)

    # 3. Distribución por Centro (CEAD / CCAV / UDR)
    por_centro: Dict[str, int] = stats.get("por_centro", {})
    if por_centro:
        pdf.seccion_titulo("3. Distribucion de Matriculas por Centro (CEAD / CCAV / UDR)")
        pdf.set_fill_color(0, 75, 135)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(100, 6.5, _texto_latin1("Centro Regional"), border=1, fill=True)
        pdf.cell(45, 6.5, _texto_latin1("Matriculas"), border=1, align="C", fill=True)
        pdf.cell(45, 6.5, _texto_latin1("Participacion (%)"), border=1, align="C", new_x="LMARGIN", new_y="NEXT", fill=True)

        pdf.set_text_color(30, 41, 59)
        pdf.set_font("Helvetica", "", 8.5)
        for cen, cant in sorted(por_centro.items(), key=lambda x: x[1], reverse=True):
            pct = (cant / total_mat * 100.0) if total_mat else 0.0
            pdf.cell(100, 6, _texto_latin1(cen)[:55], border=1)
            pdf.cell(45, 6, f"{cant:,}".replace(",", "."), border=1, align="C")
            pdf.cell(45, 6, f"{pct:.1f}%", border=1, align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)

    # 4. Distribución por Programa Académico
    por_programa: Dict[str, int] = stats.get("por_programa", {})
    if por_programa:
        pdf.seccion_titulo("4. Matriculas por Programa Academico")
        pdf.set_fill_color(0, 75, 135)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(120, 6.5, _texto_latin1("Programa Academico"), border=1, fill=True)
        pdf.cell(35, 6.5, _texto_latin1("Matriculas"), border=1, align="C", fill=True)
        pdf.cell(35, 6.5, _texto_latin1("Participacion"), border=1, align="C", new_x="LMARGIN", new_y="NEXT", fill=True)

        pdf.set_text_color(30, 41, 59)
        pdf.set_font("Helvetica", "", 8)
        for prog, cant in sorted(por_programa.items(), key=lambda x: x[1], reverse=True):
            pct = (cant / total_mat * 100.0) if total_mat else 0.0
            pdf.cell(120, 6, _texto_latin1(prog)[:68], border=1)
            pdf.cell(35, 6, f"{cant:,}".replace(",", "."), border=1, align="C")
            pdf.cell(35, 6, f"{pct:.1f}%", border=1, align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)

    # 5. Distribución por Periodo Académico (Top 15)
    por_periodo: Dict[str, int] = stats.get("por_periodo", {})
    if por_periodo:
        pdf.seccion_titulo("5. Concentracion por Periodo Academico")
        pdf.set_fill_color(0, 75, 135)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(70, 6.5, _texto_latin1("Codigo de Periodo"), border=1, fill=True)
        pdf.cell(60, 6.5, _texto_latin1("Total Matriculas"), border=1, align="C", fill=True)
        pdf.cell(60, 6.5, _texto_latin1("Participacion (%)"), border=1, align="C", new_x="LMARGIN", new_y="NEXT", fill=True)

        pdf.set_text_color(30, 41, 59)
        pdf.set_font("Helvetica", "", 8)
        for per, cant in sorted(por_periodo.items(), key=lambda x: x[1], reverse=True)[:15]:
            pct = (cant / total_mat * 100.0) if total_mat else 0.0
            pdf.cell(70, 5.8, _texto_latin1(per), border=1)
            pdf.cell(60, 5.8, f"{cant:,}".replace(",", "."), border=1, align="C")
            pdf.cell(60, 5.8, f"{pct:.1f}%", border=1, align="C", new_x="LMARGIN", new_y="NEXT")

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer.getvalue()


def generar_pdf_informe_creditos_ecisa(resumen: Dict[str, Any]) -> bytes:
    """
    Genera en memoria (`bytes`) el Informe Ejecutivo de Créditos Proyectados (`proy_zon`)
    vs Realizados (`creditos`) en 2026 por Programa (`pro_descripcion`) y Centro (`nombrecead`).
    """
    pdf = ReporteAcademicoPDF(
        titulo_reporte="INFORME COMPARATIVO: CREDITOS PROYECTADOS VS REALIZADOS 2026 (ECISA)",
        subtitulo_reporte="Cumplimiento de Metas por Programa Academico (SNIES) y Centro Regional (CEAD / CCAV / UDR)",
    )
    pdf.alias_nb_pages()
    pdf.add_page()

    tot_proy = float(resumen.get("total_proyectado", 0.0))
    tot_real = float(resumen.get("total_realizado", 0.0))
    dif = float(resumen.get("diferencia", 0.0))
    cump = float(resumen.get("cumplimiento_pct", 0.0))

    # 1. KPIs Globales
    pdf.seccion_titulo("1. Balance General de Cumplimiento de Creditos 2026")
    pdf.set_text_color(30, 41, 59)
    kpis = [
        ("Creditos Proyectados en Zona (proy_zon):", f"{tot_proy:,.0f}".replace(",", ".")),
        ("Creditos Realizados / Matriculados (creditos):", f"{tot_real:,.0f}".replace(",", ".")),
        ("Variacion / Diferencia Neta (Realizados - Proyectados):", f"{dif:+,.0f}".replace(",", ".")),
        ("Porcentaje de Cumplimiento de Meta Global:", f"{cump:.1f}%"),
    ]
    for k, v in kpis:
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(95, 6, _texto_latin1(k))
        pdf.set_font("Helvetica", "", 9)
        pdf.cell(0, 6, _texto_latin1(v), new_x="LMARGIN", new_y="NEXT")

    pdf.ln(3)

    # 2. Tabla por Centro (nombrecead)
    df_cead = resumen.get("df_por_cead")
    if df_cead is not None and not df_cead.empty:
        pdf.seccion_titulo("2. Comparativo Proyectado vs Realizado por Centro (nombrecead)")
        pdf.set_fill_color(0, 75, 135)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(72, 6.5, _texto_latin1("Nombre del CEAD / Centro"), border=1, fill=True)
        pdf.cell(30, 6.5, _texto_latin1("Proyectado"), border=1, align="R", fill=True)
        pdf.cell(30, 6.5, _texto_latin1("Realizado"), border=1, align="R", fill=True)
        pdf.cell(30, 6.5, _texto_latin1("Diferencia"), border=1, align="R", fill=True)
        pdf.cell(28, 6.5, _texto_latin1("Cumplim. (%)"), border=1, align="C", new_x="LMARGIN", new_y="NEXT", fill=True)

        pdf.set_text_color(30, 41, 59)
        pdf.set_font("Helvetica", "", 8)
        for _, r in df_cead.iterrows():
            cead = _texto_latin1(r["nombrecead"])[:38]
            proy = f"{float(r['proy_zon']):,.0f}".replace(",", ".")
            cred = f"{float(r['creditos']):,.0f}".replace(",", ".")
            d_val = f"{float(r['diferencia']):+,.0f}".replace(",", ".")
            c_pct = f"{float(r['cumplimiento_pct']):.1f}%"

            pdf.cell(72, 6, cead, border=1)
            pdf.cell(30, 6, proy, border=1, align="R")
            pdf.cell(30, 6, cred, border=1, align="R")
            pdf.cell(30, 6, d_val, border=1, align="R")
            pdf.cell(28, 6, c_pct, border=1, align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(4)

    # 3. Tabla por Programa (codigo_real_snies + pro_descripcion)
    df_prog = resumen.get("df_por_programa")
    if df_prog is not None and not df_prog.empty:
        pdf.seccion_titulo("3. Comparativo Proyectado vs Realizado por Programa (pro_descripcion)")
        pdf.set_fill_color(0, 75, 135)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(22, 6.5, _texto_latin1("SNIES"), border=1, align="C", fill=True)
        pdf.cell(74, 6.5, _texto_latin1("Programa Academico"), border=1, fill=True)
        pdf.cell(24, 6.5, _texto_latin1("Proyectado"), border=1, align="R", fill=True)
        pdf.cell(24, 6.5, _texto_latin1("Realizado"), border=1, align="R", fill=True)
        pdf.cell(24, 6.5, _texto_latin1("Diferencia"), border=1, align="R", fill=True)
        pdf.cell(22, 6.5, _texto_latin1("Cumpl. %"), border=1, align="C", new_x="LMARGIN", new_y="NEXT", fill=True)

        pdf.set_text_color(30, 41, 59)
        pdf.set_font("Helvetica", "", 7.5)
        for _, r in df_prog.iterrows():
            snies = _texto_latin1(r["codigo_real_snies"])[:12]
            prog = _texto_latin1(r["pro_descripcion"])[:42]
            proy = f"{float(r['proy_zon']):,.0f}".replace(",", ".")
            cred = f"{float(r['creditos']):,.0f}".replace(",", ".")
            d_val = f"{float(r['diferencia']):+,.0f}".replace(",", ".")
            c_pct = f"{float(r['cumplimiento_pct']):.1f}%"

            pdf.cell(22, 6, snies, border=1, align="C")
            pdf.cell(74, 6, prog, border=1)
            pdf.cell(24, 6, proy, border=1, align="R")
            pdf.cell(24, 6, cred, border=1, align="R")
            pdf.cell(24, 6, d_val, border=1, align="R")
            pdf.cell(22, 6, c_pct, border=1, align="C", new_x="LMARGIN", new_y="NEXT")

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer.getvalue()


