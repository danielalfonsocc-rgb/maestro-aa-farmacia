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
         "El umbral CDL ≥ 30 ud/día define Rotación ALTA (3 gavetas) solo para el cálculo "
         "de Cap_Max. Los umbrales de Freq_Revision usan FREQ_CDL_UMBRALES distintos."),
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
         "S (~120 cm³) → 23 cajas/gaveta  |  M (~180 cm³) → 15 cajas/gaveta  |  L (~300 cm³) → 9 cajas/gaveta",
         "Se asigna automáticamente por forma farmacéutica; se puede sobreescribir en cenabast_tallas.csv."),
        ("Rotación",
         "Intensidad de uso del medicamento en términos de gavetas asignadas. ALTA si CDL ≥ 30 "
         "ud/día (usa 3 gavetas en paralelo), NORMAL si CDL < 30 (1 gaveta).",
         None, None),
        ("Cap_Cajas",
         "Número total de cajas que caben en el espacio asignado al medicamento, "
         "considerando las gavetas según su rotación.",
         "Cap_Cajas = Base_Cajas × Gavetas_Rotacion",
         "Base_Cajas viene de la Talla o de ⌊2790 / Vol_Caja⌋ para formatos con volumen conocido."),
        ("Cap_Max (Capacidad Máxima de referencia)",
         "Número máximo de unidades que puede almacenar el espacio asignado al medicamento. "
         "Sirve como dato informativo y para activar [ALERTA_ESTRES]; NO limita el Nivel Objetivo T.",
         "Cap_Max = Cap_Cajas × Unidades_por_Caja",
         "Visible en la tabla SGLI como columna 'Cap. Máx (ud)'. El Nivel Objetivo T se calcula "
         "únicamente desde la demanda proyectada, sin techo de capacidad."),
    ]),

    ("4. Motor SGLI — Reposición Basada en Demanda", [
        ("SGLI (Sistema de Gestión Logística Integrada)",
         "Motor de cálculo de reposición de la Farmacia AA. Determina el intervalo de reposición, "
         "el nivel objetivo de stock y las acciones a tomar (traspaso o compra urgente) "
         "basándose exclusivamente en la demanda real. El modelo no tiene techo de capacidad física: "
         "el objetivo se adapta al consumo, independientemente del mueble actual.",
         None, None),
        ("IR Global (Intervalo de Reposición)",
         "Número de días hábiles entre reposiciones, calculado de forma global para todos los "
         "medicamentos a partir del Factor de Carga. Determina el Nivel Objetivo (T) y el Déficit "
         "de cada ciclo de pedido.",
         "IR = max(1, min(5, ⌊5 / Factor_Carga⌋))",
         "Con FC = 1.15 → IR = 4 días hábiles. Con FC = 1.0 → IR = 5 días hábiles."),
        ("Freq_Revision (Frecuencia de Revisión por Medicamento)",
         "Número de días hábiles entre revisiones de stock para cada medicamento individualmente. "
         "Se calcula por medicamento según su rotación histórica, criticidad y variabilidad de consumo. "
         "Determina el Stock_Minimo (ROP) per-medicamento.",
         "Bucket base CDL:  ≥150 → 1 (diario)  |  30-150 → 2  |  5-30 → 3  |  0.5-5 → 4  |  <0.5 → 5 (semanal)\n"
         "Ajuste criticidad 1-2: −1 nivel  |  Ajuste variabilidad CDL > 50%: −1 nivel\n"
         "Resultado final: max(1, min(5, bucket + ajustes))",
         "Calibrada con la distribución real del AA (404 meds con consumo): ~20% por bucket. "
         "Visible en la hoja SGLI_Estres, columna 'Freq_Revision'."),
        ("Stock_Minimo (ROP por Medicamento)",
         "Nivel mínimo de stock en Farmacia AA por debajo del cual se debe iniciar la reposición. "
         "Se calcula de forma individual para cada medicamento usando su Freq_Revision y CDL.",
         "Stock_Minimo = ⌈CDL × Freq_Revision⌉ + ⌈CDL × SS_dias⌉\n"
         "SS_dias = 1 (blindaje lunes) + 1 extra si criticidad nivel 1-2",
         "Ejemplos: Paracetamol (CDL=3695, Freq=1, nivel 1) → Stock_Min ≈ 14 780 ud. "
         "Enalapril 10mg (CDL=677, Freq=1, nivel 2) → Stock_Min ≈ 2 708 ud."),
        ("Factor de Carga (FC)",
         "Multiplicador de demanda aplicado al cálculo SGLI para simular escenarios de mayor "
         "presión asistencial. Baseline = 1.15 (15% sobre la demanda histórica).",
         None,
         "Rango típico: 1.00 (sin estrés) a 2.00 (doble de demanda). Ajustable en la pestaña SGLI de la app."),
        ("Demanda",
         "Unidades proyectadas a despachar en el intervalo de reposición, incorporando el factor de carga.",
         "Demanda = IR × CDL × Factor_Carga", None),
        ("T (Nivel Objetivo)",
         "Cantidad de unidades que debe haber en Farmacia AA al inicio de cada ciclo de reposición. "
         "Se basa únicamente en la demanda proyectada (sin techo de capacidad física), "
         "lo que permite adaptarse a muebles de mayor capacidad y a variaciones reales de consumo.",
         "T = ⌊Demanda × 1.25⌋",
         "El factor 1.25 es el margen de seguridad del 25% adicional sobre la demanda esperada."),
        ("Déficit",
         "Diferencia entre el Nivel Objetivo y el stock actual de Farmacia AA. Si es positivo, "
         "existe una necesidad de reposición.",
         "Deficit = max(0, T - Stock_Farm)", None),
        ("Margen de Seguridad (1.25)",
         "Factor de colchón del 25% adicional sobre la demanda proyectada para absorber variabilidad "
         "en el consumo. Se aplica directamente sobre la Demanda al calcular T.",
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
         "urgente por la diferencia.",
         "0 < Stock_Bod < Deficit → TRASPASAR [Stock_Bod] + COMPRA URGENTE [Deficit − Stock_Bod]\n"
         "Stock_Bod = 0        → SIN STOCK BODEGA + COMPRA URGENTE [Deficit]",
         None),
        ("[ALERTA_ESTRES]",
         "Indicador que se activa cuando el ROP (punto de reorden) supera la Cap_Max física "
         "estimada de la gaveta, señalando que el mueble actual puede resultar insuficiente "
         "incluso para el stock mínimo de seguridad.",
         "Se activa si ROP > Cap_Max",
         "Cap_Max es informativo; no limita T. La alerta avisa que hace falta ampliar el espacio físico."),
    ]),

    ("6. Reposición por Días Hábiles", [
        ("Días Hábiles",
         "Días de la semana en que opera la farmacia (lunes a viernes), excluyendo feriados "
         "nacionales y regionales de Chile. El archivo feriados_chile.csv define el calendario.",
         None,
         "La farmacia NO dispensa fines de semana (< 0.2% del consumo histórico). "
         "El blindaje es la reapertura del lunes, no el consumo sáb/dom."),
        ("IR en Reposición por Días Hábiles",
         "El módulo reposicion_dias_habiles.py usa la Freq_Revision per-medicamento "
         "(de la hoja SGLI_Estres) para calcular el ROP de cada medicamento. Si Freq_Revision "
         "no está disponible, recurre al IR global como fallback.",
         "IR_efectivo = Freq_Revision (si disponible) o Dias_Reposicion_IR (IR global)", None),
        ("ROP (Re-Order Point / Punto de Reorden)",
         "Nivel de stock de Farmacia en el que se debe iniciar la reposición. "
         "Desde la incorporación de Freq_Revision, el ROP es per-medicamento y equivale "
         "al Stock_Minimo de la hoja SGLI_Estres.",
         "ROP = ⌈CDL × Freq_Revision⌉ + ⌈CDL × SS_dias⌉",
         "SS_dias = 1 (reapertura lunes) + 1 extra si criticidad 1-2. "
         "El ROP y el Stock_Minimo de SGLI_Estres usan la misma fórmula."),
        ("SS (Stock de Seguridad)",
         "Colchón de unidades destinado a blindar la reapertura del lunes y absorber variaciones "
         "imprevistas de demanda.",
         "SS = ⌈CDL × SS_dias⌉   con SS_dias = 1 (normal) o 2 (criticidad 1-2)", None),
        ("Próxima Reposición",
         "Fecha del primer día hábil en que el stock proyectado cae al ROP, avanzando por el "
         "calendario de días hábiles (saltando fines de semana y feriados). "
         "Calculada por reposicion_dias_habiles.py.",
         None,
         "La columna 'Cubre_Cierre' indica si el stock actual aguanta hasta después del próximo cierre."),
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
         "en el mismo período sin justificación clínica. "
         "Detectada por auditoria_duplicados_profunda.py (histórico) y por agente_duplicados.py (operacional).",
         None, None),
        ("ISP (Recetas Cheque)",
         "Control de recetas que verifica el cumplimiento de la normativa del Instituto de Salud "
         "Pública para medicamentos psicotrópicos y estupefacientes. Módulo: recetas_cheque.py.",
         None, None),
        ("Agente de Duplicados (agente_duplicados.py)",
         "Agente IA (Claude Haiku) que analiza las recetas del día, detecta prescripciones "
         "duplicadas y genera un reporte Excel con razonamiento clínico. Los RUTs nunca salen "
         "del proceso local: se envían IDs anónimos (hash SHA-256) a la API.",
         "py agente_duplicados.py --fecha YYYY-MM-DD --ventana 60",
         "A diferencia de auditoria_duplicados_profunda.py (histórico completo), este agente es operacional: "
         "foco en las recetas de hoy y sus antecedentes recientes."),
        ("Auditoría de Medicamento (auditoria_medicamento.py)",
         "Script genérico que audita CMP dispensado, prescriptores, diagnósticos y duplicidad "
         "de prescripción para cualquier medicamento del AA.",
         "py auditoria_medicamento.py --contiene METFORMINA\n"
         "py auditoria_medicamento.py --contiene EMPAGLIFLOZINA --dosis 10",
         None),
    ]),

    ("8. Centinela de Medicamentos (Campaña Invierno)", [
        ("Centinela",
         "Seguimiento semanal de un grupo reducido de medicamentos críticos durante la campaña "
         "de invierno (aumento de IRA/ERA). Genera un informe PDF estructurado para reportar "
         "al MINSAL con proyecciones de consumo y análisis de variaciones.",
         None, None),
        ("centinela_reporte.py",
         "Script que genera el reporte centinela semanal. Lee los datos de stock y recetas, "
         "calcula proyecciones inteligentes y exporta el PDF listo para adjuntar al MINSAL.",
         "py centinela_reporte.py",
         "Los informes se guardan como centinela_SXX.pdf (número de semana epidemiológica)."),
        ("Semana Epidemiológica (S##)",
         "Numeración oficial de semanas del año epidemiológico chileno. El reporte centinela "
         "se identifica por su número de semana (ej. S25, S52).",
         None, None),
    ]),

    ("9. Módulos Compartidos y Utilidades", [
        ("utils_aa.py",
         "Módulo de utilidades compartidas del proyecto. Contiene norm_erp (normalización de "
         "nombres), HOMOLOGACION (tabla canónica de 20 entradas), cargar_recetas_csv y "
         "setup_stdout. Regla: nuevas homologaciones se agregan SOLO aquí, nunca en scripts individuales.",
         "from utils_aa import norm_erp, HOMOLOGACION, cargar_recetas_csv",
         "Centraliza lo que antes estaba duplicado en auditoria_prescripcion.py y agente_duplicados.py."),
        ("norm_erp(s)",
         "Función de utils_aa.py que normaliza nombres de medicamentos: NFD + elimina diacríticos "
         "+ colapsa espacios múltiples + convierte a mayúsculas. Garantiza comparaciones robustas "
         "entre el ERP (SSASur) y el maestro.",
         "norm_erp('Ácido Fólico  1 mg') → 'ACIDO FOLICO 1 MG'", None),
        ("HOMOLOGACION",
         "Tabla canónica de equivalencias de nombres: mapea el nombre como viene del ERP al "
         "nombre canónico del maestro. Contiene 20 entradas. Fuente única: utils_aa.py.",
         None,
         "Ejemplo: 'TRAZODONA CM 100 MG' → 'TRAZODONA CM  100 MG'"),
        ("cargar_recetas_csv(work_dir, cols, solo_ultimo)",
         "Función de utils_aa.py que carga y deduplica todos los CSV de recetas SSASur "
         "(informe_completo_recetas*.csv) de la carpeta, eliminando filas duplicadas por "
         "'ID Receta Detalle'. El flag solo_ultimo carga solo el archivo más reciente.",
         None, None),
        ("aa_colors.py",
         "Módulo de paleta de colores compartida para Excel e impresión. Define VINO, TEAL, "
         "ROJO, NARANJA, AMBAR, VERDE, GRIS_CLR y helpers soften/darken para bandas tenues "
         "de impresión económica.",
         None,
         "Paleta tinta-económica: crítico = rosa pastel + texto vino. Sin bandas oscuras."),
    ]),

    ("10. Fuentes de Datos y Archivos", [
        ("SSASur",
         "Servicio de Salud Araucanía Sur. Plataforma institucional desde donde se descarga "
         "el stock (reporte_de_stock) y las recetas (informe_completo_recetas) vía AUTO_SSASUR.py. "
         "La descarga de recetas exige entrar por la tarjeta RECETA (proyecto 629 vía SSO). "
         "RCE/deep-link devuelve 0 filas. Bloques máximos de 30 días.",
         None, None),
        ("Consolidado_AA_MAESTRO.xlsx",
         "Archivo Excel de salida principal con 14 hojas. Contiene todas las métricas del "
         "sistema: stock, demanda, SGLI, alertas y reposición. Generado por maestro_aa.py.",
         None, None),
        ("cenabast_tallas.csv",
         "Archivo de overrides por medicamento: permite especificar manualmente la Talla, "
         "volumen de caja y unidades por caja cuando los valores por defecto no son precisos.",
         "Columnas: Medicamento ; Talla ; Vol_Caja_cm3 ; Unidades_Caja", None),
        ("cenabast_intermediacion.csv",
         "Archivo con datos de intermediación CENABAST por medicamento (precios, licitaciones). "
         "Complementa el cálculo de evaluación de compra ágil.",
         None, None),
        ("feriados_chile.csv",
         "Calendario de feriados nacionales y regionales de Chile. Usado por "
         "reposicion_dias_habiles.py para calcular días hábiles reales. "
         "Debe mantenerse actualizado anualmente.",
         "Columnas: Fecha (ISO) ; Nombre ; Confianza", None),
        ("SGLI_Estres (hoja del Consolidado)",
         "Hoja del Consolidado que persiste el resultado del motor SGLI con el Factor_Carga "
         "baseline. La pestaña 'SGLI / Capacidad' de la app lo recalcula en vivo. "
         "Columnas clave: Freq_Revision, Stock_Minimo, Cap_Max, Nivel_Objetivo_T, Deficit.",
         None, None),
        ("informe_completo_recetas*.csv",
         "Archivos de recetas descargados de SSASur (uno por bloque de 30 días). "
         "Pueden contener RUT de pacientes — datos sujetos a Ley 19.628. "
         "Nunca pegar filas con RUT en el chat; anonimizar antes de depurar.",
         None, None),
    ]),

    ("11. Gestión Territorial (GT)", [
        ("GT (Gestión Territorial)",
         "Proceso de distribución de medicamentos desde el Hospital de Pitrufquén hacia los "
         "establecimientos de la red Araucanía Sur (CESFAM, postas, hospitales del nodo).",
         None, None),
        ("Guía de Tratamiento (GT)",
         "Documento que acompaña el traslado de medicamentos hacia un establecimiento de la red. "
         "cruce_gt.py cruza las guías con el Consolidado para detectar inconsistencias.",
         None, None),
        ("cruce_gt.py",
         "Script que cruza las Guías de Tratamiento descargadas de SSASur con el Consolidado "
         "para detectar diferencias entre lo enviado y el stock registrado.",
         "py cruce_gt.py", None),
    ]),

    ("12. Constantes del Sistema", [
        ("FACTOR_CARGA_DEFAULT (1.15)",
         "Valor por defecto del factor de carga: 15% de margen sobre la demanda histórica media. "
         "Usado en todos los cálculos SGLI baseline.",
         None, None),
        ("IR_MAX (5)",
         "Intervalo de reposición máximo: 5 días hábiles (semana laboral completa). "
         "Ningún medicamento se repone con menor frecuencia que esto.",
         None, None),
        ("MARGEN_SEGURIDAD (1.25)",
         "Factor de colchón del 25% adicional sobre la demanda proyectada. "
         "Se aplica al calcular el Nivel Objetivo T.",
         None, None),
        ("VOLUMEN_GAVETA_CM3 (2 790 cm³)",
         "Volumen de referencia de una gaveta estándar de Farmacia AA. "
         "Se usa para estimar cuántas cajas caben cuando se conoce el volumen de caja.",
         None, None),
        ("UMBRAL_ALTA_ROTACION_CDL (30 ud/día)",
         "Umbral de CDL para clasificar Rotación ALTA (3 gavetas). "
         "Solo afecta Cap_Max; no determina Freq_Revision.",
         None, None),
        ("FREQ_CDL_UMBRALES",
         "Tabla de umbrales CDL que define la Freq_Revision por medicamento. "
         "Calibrada con la distribución empírica real del AA para que cada bucket "
         "contenga aproximadamente el 20% de los medicamentos con consumo.",
         "CDL >= 150 → 1 (diario, ~19%)  |  CDL 30-150 → 2 (~20%)  |  CDL 5-30 → 3 (~20%)\n"
         "CDL 0.5-5 → 4 (cada 4 días, ~19%)  |  CDL < 0.5 → 5 (semanal, ~22%)",
         "Separado de UMBRAL_ALTA_ROTACION_CDL=30, que es solo para gavetas."),
        ("BUFFER_FINDE_DIAS (1)",
         "Días de stock de seguridad para blindar la reapertura del lunes. "
         "Valor = 1 porque la demanda sáb/dom es ~0% (medido sobre Fecha_Entrega real).",
         None, None),
        ("EXTRA_SS_CRITICOS (1)",
         "Día adicional de SS para medicamentos con criticidad nivel 1-2. "
         "Se suma al BUFFER_FINDE_DIAS para los críticos (no pueden quebrar stock).",
         None, None),
        ("App Pedidos (app_pedidos.py)",
         "Dashboard Streamlit que centraliza todas las vistas operativas: pedidos, faltantes, "
         "SGLI, reposición por días hábiles, auditorías y alertas. Corre en el puerto 8501.",
         "py -m streamlit run app_pedidos.py --server.port 8501", None),
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

    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("GLOSARIO DE TERMINOLOGÍA", titulo_doc))
    story.append(Paragraph("Sistema Maestro AA · Farmacia Atención Abierta", subtitulo_doc))
    story.append(Paragraph(
        f"Hospital de Pitrufquén — Servicio de Salud Araucanía Sur  |  {datetime.date.today().strftime('%d/%m/%Y')}",
        fecha_style,
    ))
    story.append(HRFlowable(width="100%", thickness=2, color=VINO, spaceAfter=8))

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
