"""Genera Glosario_Maestro_AA.pdf en la carpeta Farmacia AA del escritorio."""
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether,
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os, datetime

DEST = r"C:\Users\danie\OneDrive\Desktop\Farmacia AA\Glosario_Maestro_AA.pdf"

# ── Paleta ────────────────────────────────────────────────────────────────────
VINO      = colors.HexColor("#7C2D12")
VINO_SUAVE= colors.HexColor("#B45309")
ROSA      = colors.HexColor("#FEF2F2")
CREMA     = colors.HexColor("#FFFBEB")
GRIS_LINE = colors.HexColor("#E5E7EB")
NEGRO     = colors.HexColor("#111827")
GRIS_TEXT = colors.HexColor("#374151")

# ── Estilos ───────────────────────────────────────────────────────────────────
styles = getSampleStyleSheet()

titulo_doc = ParagraphStyle(
    "TituloDoc", parent=styles["Title"],
    textColor=VINO, fontSize=22, spaceAfter=4, alignment=TA_CENTER,
    fontName="Helvetica-Bold",
)
subtitulo_doc = ParagraphStyle(
    "SubtituloDoc", parent=styles["Normal"],
    textColor=VINO_SUAVE, fontSize=11, spaceAfter=2, alignment=TA_CENTER,
    fontName="Helvetica",
)
fecha_style = ParagraphStyle(
    "FechaStyle", parent=styles["Normal"],
    textColor=GRIS_TEXT, fontSize=9, spaceAfter=16, alignment=TA_CENTER,
    fontName="Helvetica-Oblique",
)
seccion = ParagraphStyle(
    "Seccion", parent=styles["Heading2"],
    textColor=VINO, fontSize=13, spaceBefore=18, spaceAfter=6,
    fontName="Helvetica-Bold", borderPad=0,
)
termino = ParagraphStyle(
    "Termino", parent=styles["Normal"],
    textColor=VINO, fontSize=10, fontName="Helvetica-Bold",
    spaceBefore=0, spaceAfter=1,
)
definicion = ParagraphStyle(
    "Definicion", parent=styles["Normal"],
    textColor=GRIS_TEXT, fontSize=9.5, fontName="Helvetica",
    spaceBefore=0, spaceAfter=4, alignment=TA_JUSTIFY, leading=13,
)
formula_style = ParagraphStyle(
    "Formula", parent=styles["Normal"],
    textColor=NEGRO, fontSize=9, fontName="Courier",
    spaceBefore=1, spaceAfter=4, leftIndent=12,
    backColor=CREMA, borderPad=4,
)
nota_style = ParagraphStyle(
    "Nota", parent=styles["Normal"],
    textColor=GRIS_TEXT, fontSize=8.5, fontName="Helvetica-Oblique",
    spaceBefore=0, spaceAfter=6, leftIndent=12,
)

# ── Contenido del glosario ────────────────────────────────────────────────────
GLOSARIO = [
    # (seccion, [(termino, definicion, formula_o_None, nota_o_None), ...])

    ("1. Identificadores de Medicamentos y Grupos", [
        ("AA (Atención Abierta)",
         "Modalidad de dispensación ambulatoria del Hospital de Pitrufquén (SSASur). "
         "El universo Maestro AA comprende 378 medicamentos gestionados bajo esta modalidad.",
         None, None),
        ("Medicamento AA",
         "Fármaco inscrito en el listado de 378 medicamentos de la Farmacia de Atención Abierta "
         "del Hospital de Pitrufquén. Es la unidad mínima de análisis del sistema.",
         None, None),
        ("Criticidad",
         "Clasificación de prioridad de un medicamento según su relevancia clínica y riesgo de "
         "desabastecimiento. El sistema maneja 5 niveles.",
         "1-CRITICO | 2-ALTO | 3-MODERADO | 4-BAJO | 5-OK",
         "Aparece en la columna 'Criticidad' del Consolidado y determina el orden de las acciones logísticas."),
        ("Forma Farmacéutica",
         "Presentación física del medicamento. El motor SGLI la detecta automáticamente desde el "
         "nombre para asignar la Talla de gaveta y las unidades por caja por defecto.",
         None,
         "Categorías: SOLIDO_ORAL · INYECTABLE · LIQUIDO_ORAL · INHALADOR · TOPICO · "
         "INSULINA · OFTALMICO · OVULO · SOBRE · PARCHE · OTRO"),
    ]),

    ("2. Stock y Flujo de Medicamentos", [
        ("Stock Farmacia AA (Stock_Farm)",
         "Cantidad de unidades disponibles físicamente en la Farmacia de Atención Abierta al "
         "momento de la consulta. Fuente: sistema SSASur.",
         None, None),
        ("Stock Bodega AA (Stock_Bod)",
         "Cantidad de unidades disponibles en la Bodega del Hospital que pueden trasladarse "
         "a Farmacia AA. Fuente: sistema SSASur.",
         None, None),
        ("CDL (Consumo Diario Laborable)",
         "Promedio de unidades dispensadas por día hábil. Se calcula dividiendo el consumo "
         "del período entre los días hábiles del mismo período (excluye feriados y fines de semana).",
         "CDL = Consumo_Periodo / Dias_Habiles_Periodo",
         "Umbral físico de gaveta: CDL ≥ 30 ud/día → Rotación ALTA (3 gavetas); si no, NORMAL (1 gaveta). "
         "Los umbrales de Freq_Revision son distintos y se calibraron con la distribución real del AA (ver sección 4)."),
        ("Consumo Mensual Promedio (CMP)",
         "Promedio de unidades dispensadas por mes calendario. Aparece en varias hojas del "
         "Consolidado como referencia de demanda mensual.",
         None, None),
        ("Factor Empaque (Factor_Empaque)",
         "Número de unidades por caja según CENABAST. Cuando se conoce, reemplaza al promedio "
         "por forma farmacéutica en el cálculo de capacidad de gaveta.",
         None, None),
    ]),

    ("3. Capacidad Física del Almacenamiento", [
        ("Gaveta",
         "Compartimento físico del sistema de almacenamiento de la Farmacia AA donde se guardan "
         "las cajas de un medicamento. Volumen estándar de referencia: 2 790 cm³.",
         None, None),
        ("Talla (S / M / L)",
         "Clasificación de la caja del medicamento según su volumen típico, que determina cuántas "
         "cajas caben en una gaveta estándar.",
         "S (~120 cm3) → 23 cajas | M (~180 cm3) → 15 cajas | L (~300 cm3) → 9 cajas",
         "Se asigna por forma farmacéutica; se puede sobreescribir en cenabast_tallas.csv."),
        ("Rotación",
         "Intensidad de uso del medicamento en términos de gavetas asignadas. ALTA si CDL ≥ 30 "
         "ud/día (usa 3 gavetas en paralelo), NORMAL si CDL < 30 (1 gaveta).",
         None, None),
        ("Cap_Cajas",
         "Número total de cajas que caben en el espacio asignado al medicamento, "
         "considerando las gavetas según su rotación.",
         "Cap_Cajas = Base_Cajas × Gavetas_Rotacion",
         "Base_Cajas viene de la Talla o de ⌊2790 / Vol_Caja⌋ para formatos externos."),
        ("Cap_Max (Capacidad Máxima)",
         "Número máximo de unidades que puede almacenar el espacio asignado al medicamento. "
         "Sirve como referencia de capacidad; desde la versión actual no limita el Nivel Objetivo.",
         "Cap_Max = Cap_Cajas × Unidades_por_Caja",
         "Visible en la tabla SGLI como columna 'Cap. Máx (ud)'."),
    ]),

    ("4. Motor SGLI — Reposición", [
        ("SGLI (Sistema de Gestión Logística Integrada)",
         "Motor de cálculo de reposición de la Farmacia AA. Determina el intervalo de reposición, "
         "el nivel objetivo de stock y las acciones a tomar (traspaso o compra urgente), "
         "basándose en la demanda real sin restricción de capacidad física.",
         None, None),
        ("IR Global (Intervalo de Reposición)",
         "Número de días hábiles entre reposiciones, calculado de forma global para todos los "
         "medicamentos a partir del Factor de Carga. Determina el Nivel Objetivo (T) y el Déficit "
         "de cada ciclo de pedido. No confundir con Freq_Revision, que es por medicamento.",
         "IR_global = max(1, min(5, ⌊5 / Factor_Carga⌋))",
         "Con FC = 1.15 → IR_global = 4 días. Con FC = 1.0 → IR_global = 5 días."),
        ("Freq_Revision (Frecuencia de Revisión por Medicamento)",
         "Número de días hábiles entre revisiones de stock para cada medicamento individualmente. "
         "A diferencia del IR global (que depende del Factor de Carga), la Freq_Revision depende "
         "de la rotación histórica del medicamento, su criticidad y la variabilidad de su consumo. "
         "Determina también el Stock_Minimo (ROP) per-medicamento.",
         "Bucket base CDL:  ≥150 → 1 (diario)  |  30-150 → 2  |  5-30 → 3  |  0.5-5 → 4  |  <0.5 → 5 (semanal)\n"
         "Ajuste criticidad 1-2: −1 nivel  |  Ajuste variabilidad CDL > 50 %: −1 nivel\n"
         "Resultado final: max(1, min(5, bucket + ajustes))",
         "Umbrales calibrados con la distribución real del AA (404 meds con consumo): "
         "cada bucket agrupa ~20 % de los medicamentos. "
         "Visible en la hoja SGLI_Estres, columna 'Freq_Revision'."),
        ("Stock_Minimo (ROP por Medicamento)",
         "Nivel mínimo de stock en Farmacia AA por debajo del cual se debe iniciar la reposición. "
         "Se calcula de forma individual para cada medicamento usando su Freq_Revision y CDL. "
         "Reemplaza al ROP global del módulo de reposición cuando está disponible en SGLI_Estres.",
         "Stock_Minimo = ⌈CDL × Freq_Revision⌉ + ⌈CDL × SS_dias⌉\n"
         "SS_dias = 1 (blindaje lunes) + 1 extra si criticidad nivel 1-2",
         "Ejemplos: Paracetamol (CDL=3695, Freq=1, nivel 1) → Stock_Min=14780 ud. "
         "Enalapril 10mg (CDL=677, Freq=1, nivel 2) → Stock_Min=2708 ud."),
        ("Factor de Carga (FC)",
         "Multiplicador de demanda aplicado al cálculo SGLI para simular escenarios de mayor "
         "presión asistencial. Baseline = 1.15 (15 % sobre la demanda histórica).",
         None,
         "Rango típico: 1.00 (sin estrés) a 2.00 (doble de demanda). Ajustable en la pestaña SGLI de la app."),
        ("Demanda",
         "Unidades proyectadas a despachar en el intervalo de reposición, incorporando el factor "
         "de carga. Es la base para calcular el Nivel Objetivo.",
         "Demanda = IR × CDL × Factor_Carga", None),
        ("T (Nivel Objetivo)",
         "Cantidad de unidades que debe haber en Farmacia AA al inicio de cada ciclo de "
         "reposición. Desde la versión actual, el objetivo se basa únicamente en la demanda "
         "(sin techo de capacidad física), permitiendo adaptarse a muebles de mayor capacidad.",
         "T = ⌊Demanda × 1.25⌋",
         "El 1.25 es el margen de seguridad: el 25 % adicional actúa como colchón ante variaciones de consumo."),
        ("Déficit",
         "Diferencia entre el Nivel Objetivo y el stock actual de Farmacia AA. Si es positivo, "
         "existe una necesidad de reposición.",
         "Deficit = max(0, T - Stock_Farm)", None),
        ("Margen de Seguridad (1.25)",
         "Factor de colchón aplicado sobre la demanda proyectada para absorber variabilidad en el "
         "consumo. Corresponde al 25 % adicional sobre la demanda esperada.",
         None, None),
    ]),

    ("5. Acciones Logísticas", [
        ("Acción 1 — Traspaso desde Bodega",
         "Cuando hay déficit y existe stock en Bodega AA, se indica traspasar la cantidad necesaria "
         "(o disponible, si no alcanza) de Bodega a Farmacia.",
         "Stock_Bod ≥ Deficit → TRASPASAR [Deficit] DESDE BODEGA",
         None),
        ("Acción 2 — Compra Urgente / Externa",
         "Cuando el stock de Bodega no cubre el déficit total, se genera una orden de compra "
         "urgente por la diferencia. Si Bodega = 0, toda la necesidad es compra urgente.",
         "0 < Stock_Bod < Deficit → TRASPASAR [Stock_Bod] + COMPRA URGENTE [Deficit − Stock_Bod]\n"
         "Stock_Bod = 0 → SIN STOCK BODEGA + COMPRA URGENTE [Deficit]",
         None),
        ("[ALERTA_ESTRES]",
         "Indicador que se activa cuando el ROP (punto de reorden) supera la capacidad física "
         "estimada de la gaveta, señalando que el mueble actual puede resultar insuficiente "
         "incluso para el stock mínimo de seguridad.",
         "Se activa si ROP > Cap_Max",
         "ROP = CDL × IR + SS  (ver sección Reposición por Días Hábiles)."),
    ]),

    ("6. Reposición por Días Hábiles", [
        ("Días Hábiles",
         "Días de la semana en que opera la farmacia (lunes a viernes), excluyendo feriados "
         "nacionales y regionales de Chile. El archivo feriados_chile.csv define el calendario.",
         None,
         "La farmacia NO dispensa fines de semana (< 0.2 % del consumo histórico). "
         "El blindaje es la reapertura del lunes."),
        ("IR usado en Reposición por Días Hábiles",
         "El módulo reposicion_dias_habiles.py utiliza la Freq_Revision per-medicamento "
         "(de la hoja SGLI_Estres) para calcular el ROP de cada medicamento. Si Freq_Revision "
         "no está disponible, recurre al IR global como fallback.",
         "IR_efectivo = Freq_Revision (si disponible) o Dias_Reposicion_IR (IR global)", None),
        ("ROP (Re-Order Point / Punto de Reorden)",
         "Nivel de stock de Farmacia en el que se debe iniciar la reposición para no quedar "
         "desabastecido antes de que llegue el siguiente traspaso. Desde la incorporación de "
         "Freq_Revision, el ROP es per-medicamento (= Stock_Minimo en SGLI_Estres).",
         "ROP = ⌈CDL × Freq_Revision⌉ + ⌈CDL × SS_dias⌉",
         "SS_dias = 1 (reapertura lunes) + 1 extra si criticidad 1-2. "
         "El Stock_Minimo de SGLI_Estres y el ROP de Reposicion_DiasHabiles usan la misma fórmula."),
        ("SS (Stock de Seguridad)",
         "Colchón de unidades destinado a blindar la reapertura del lunes y absorber variaciones "
         "imprevistas de demanda. La farmacia no dispensa fines de semana (< 0.2 % del consumo "
         "histórico), por lo que el blindaje real es cubrir la demanda del primer día hábil tras "
         "el cierre, no el consumo del fin de semana.",
         "SS = ⌈CDL × SS_dias⌉   con SS_dias = 1 (normal) o 2 (criticidad 1-2)", None),
    ]),

    ("7. Auditorías de Recetas", [
        ("Receta",
         "Documento que autoriza la dispensación de un medicamento. En SSASur, las recetas "
         "crónicas anuales generan un número de receta por cuota mensual.",
         None,
         "Una receta anual puede aparecer hasta 12 veces con distinto N° de Receta. "
         "Colapsar por EVENTO (RUN + med + médico + fecha) para evitar duplicidad ×12."),
        ("EVENTO",
         "Unidad analítica para auditorías de recetas: combinación de RUN del paciente, "
         "medicamento, médico prescriptor y fecha. Evita inflar la cuenta por cuotas mensuales.",
         "EVENTO = RUN + Medicamento + Médico + Fecha_Receta", None),
        ("Duplicidad de Recetas",
         "Situación en que un mismo medicamento es dispensado más de una vez al mismo paciente "
         "en el mismo período sin justificación clínica. Detectada por recetas_duplicadas.py.",
         None, None),
        ("ISP (Recetas Cheque)",
         "Control de recetas que verifica el cumplimiento de la normativa del Instituto de Salud "
         "Pública para medicamentos psicotrópicos y estupefacientes. Módulo: recetas_cheque.py.",
         None, None),
    ]),

    ("8. Fuentes de Datos y Archivos", [
        ("SSASur",
         "Servicio de Salud Araucanía Sur. Plataforma institucional desde donde se descarga "
         "el stock (reporte_de_stock) y las recetas (informe_completo_recetas) vía AUTO_SSASUR.py.",
         None, None),
        ("Consolidado_AA_MAESTRO.xlsx",
         "Archivo Excel de salida principal con 14 hojas. Contiene todas las métricas del "
         "sistema: stock, demanda, SGLI, alertas y reposición.",
         None, None),
        ("cenabast_tallas.csv",
         "Archivo de overrides por medicamento: permite especificar manualmente la Talla, "
         "volumen de caja y unidades por caja cuando los valores por defecto no son precisos.",
         "Columnas: Medicamento ; Talla ; Vol_Caja_cm3 ; Unidades_Caja", None),
        ("feriados_chile.csv",
         "Calendario de feriados nacionales y regionales de Chile. Usado por "
         "reposicion_dias_habiles.py para calcular días hábiles reales. Debe mantenerse actualizado.",
         None, None),
        ("SGLI_Estres (hoja del Consolidado)",
         "Hoja del Consolidado que persiste el resultado del motor SGLI con el Factor_Carga "
         "baseline. La pestaña 'SGLI / Capacidad' de la app lo recalcula en vivo.",
         None, None),
    ]),

    ("9. Indicadores de Gestión Territorial (GT)", [
        ("GT (Gestión Territorial)",
         "Proceso de distribución de medicamentos desde el Hospital de Pitrufquén hacia los "
         "establecimientos de la red Araucanía Sur (CESFAM, postas, hospitales del nodo).",
         None, None),
        ("Guía de Tratamiento (GT)",
         "Documento que acompaña el traslado de medicamentos hacia un establecimiento de la red. "
         "cruce_gt.py cruza las guías con el Consolidado para detectar inconsistencias.",
         None, None),
    ]),

    ("10. Términos Generales del Sistema", [
        ("Factor_Carga_Default (1.15)",
         "Valor por defecto del factor de carga usado en todos los cálculos SGLI. "
         "Representa un 15 % de margen sobre la demanda histórica media.",
         None, None),
        ("IR_MAX (5)",
         "Intervalo de reposición máximo permitido en el sistema: 5 días hábiles (una semana "
         "laboral completa). Ningún medicamento se repone con menor frecuencia que esto.",
         None, None),
        ("VOLUMEN_GAVETA_CM3 (2 790 cm³)",
         "Volumen de referencia de una gaveta estándar del sistema de almacenamiento de "
         "Farmacia AA. Se usa para estimar cuántas cajas caben cuando se conoce el volumen de caja.",
         None, None),
        ("UMBRAL_ALTA_ROTACION_CDL (30 ud/día)",
         "Umbral de CDL a partir del cual un medicamento se clasifica como Rotación ALTA "
         "y se le asignan 3 gavetas en lugar de 1. Exclusivo para el cálculo de Cap_Max. "
         "No determina la Freq_Revision (que usa FREQ_CDL_UMBRALES distintos).",
         None, None),
        ("FREQ_CDL_UMBRALES",
         "Tabla de umbrales CDL que define la frecuencia de revisión (Freq_Revision) "
         "por medicamento. Calibrada con la distribución empírica real del AA para que "
         "cada bucket contenga aproximadamente el 20 % de los medicamentos con consumo.",
         "CDL >= 150 → 1 (diario, ~19 %)  |  CDL 30-150 → 2 (~20 %)  |  CDL 5-30 → 3 (~20 %)\n"
         "CDL 0.5-5 → 4 (cada 4 días, ~19 %)  |  CDL < 0.5 → 5 (semanal, ~22 %)",
         "Distinto de UMBRAL_ALTA_ROTACION_CDL=30, que es solo para gavetas. "
         "Los ajustes de criticidad y variabilidad se aplican sobre estos buckets base."),
        ("App Pedidos (app_pedidos.py)",
         "Dashboard Streamlit que centraliza todas las vistas operativas: pedidos, faltantes, "
         "SGLI, reposición por días hábiles, auditorías y alertas. Corre en el puerto 8501.",
         None, None),
    ]),
]

# ── Construcción del documento ────────────────────────────────────────────────
def build():
    doc = SimpleDocTemplate(
        DEST,
        pagesize=letter,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2.2*cm, bottomMargin=2.2*cm,
        title="Glosario Maestro AA — Farmacia Atención Abierta",
        author="Hospital de Pitrufquén · SSASur",
        subject="Terminología del Sistema de Gestión Logística",
    )

    story = []

    # Encabezado
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("GLOSARIO DE TERMINOLOGÍA", titulo_doc))
    story.append(Paragraph("Sistema Maestro AA · Farmacia Atención Abierta", subtitulo_doc))
    story.append(Paragraph(
        f"Hospital de Pitrufquén — Servicio de Salud Araucanía Sur  |  {datetime.date.today().strftime('%d/%m/%Y')}",
        fecha_style,
    ))
    story.append(HRFlowable(width="100%", thickness=2, color=VINO, spaceAfter=8))

    # Intro breve
    story.append(Paragraph(
        "Este glosario define los términos técnicos utilizados en el sistema Maestro AA para la "
        "gestión logística de medicamentos de Atención Abierta. Incluye indicadores, fórmulas, "
        "archivos de datos y módulos del sistema.",
        ParagraphStyle("Intro", parent=styles["Normal"], fontSize=9.5, textColor=GRIS_TEXT,
                       fontName="Helvetica-Oblique", spaceAfter=10, alignment=TA_JUSTIFY),
    ))

    for nombre_sec, terminos in GLOSARIO:
        story.append(Paragraph(nombre_sec, seccion))
        story.append(HRFlowable(width="100%", thickness=0.5, color=GRIS_LINE, spaceAfter=4))

        for t, d, f, n in terminos:
            bloque = []
            bloque.append(Paragraph(t, termino))
            bloque.append(Paragraph(d, definicion))
            if f:
                bloque.append(Paragraph(f, formula_style))
            if n:
                bloque.append(Paragraph(f"→ {n}", nota_style))
            story.append(KeepTogether(bloque))

    # Pie de página
    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=1, color=VINO, spaceBefore=4, spaceAfter=4))
    story.append(Paragraph(
        "Generado automáticamente por el sistema Maestro AA · Hospital de Pitrufquén",
        ParagraphStyle("Pie", parent=styles["Normal"], fontSize=8, textColor=GRIS_TEXT,
                       fontName="Helvetica-Oblique", alignment=TA_CENTER),
    ))

    doc.build(story)
    print(f"PDF generado: {DEST}")

if __name__ == "__main__":
    build()
