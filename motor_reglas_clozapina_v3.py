"""
Motor de reglas V3 — Seguimiento hematológico Clozapina
Base normativa: Norma Técnica para el Uso de Clozapina, MINSAL Chile (oct-2018)
Hospital Pitrufquén (SSAS) — atención abierta, clozapina 100 mg

V3: incorpora RAE (Recuento Absoluto de Eosinófilos) — campo que la plataforma
MINSAL despliega junto al RAN al ingresar el hemograma. Clasificación de
eosinofilia según Norma Técnica MINSAL 2018, Anexo 2.
"""

from dataclasses import dataclass, field
from datetime import date
from typing import Optional, Literal
from enum import Enum


# ============================================================
# ENUMS Y CONSTANTES NORMATIVAS
# ============================================================

class FaseNorma(str, Enum):
    SEMANAL = "semanal"
    QUINCENAL = "quincenal"
    MENSUAL = "mensual"


class AccionPlataforma(str, Enum):
    INGRESAR = "INGRESAR_HEMOGRAMA"
    ALERTAR_NUEVO = "ALERTAR_PACIENTE_NO_REGISTRADO"
    NO_INGRESAR_DATOS_INCOMPLETOS = "NO_INGRESAR_DOSIS_FASE_INCOMPLETA"
    NO_INSISTIR_BIMENSUAL = "NO_INSISTIR_INGRESO_BIMENSUAL_VIGENTE"
    ALERTAR_SIN_HEMOGRAMA = "ALERTAR_SIN_HEMOGRAMA_EN_HCE"


class NivelAlarmaRAN(str, Enum):
    SIN_ALARMA = "SIN_ALARMA"
    ALARMA_1_LEVE = "ALARMA_1_NEUTROPENIA_LEVE"           # 1000-1499
    ALARMA_2_MODERADA = "ALARMA_2_NEUTROPENIA_MODERADA"   # 500-999
    ALARMA_3_GRAVE = "ALARMA_3_NEUTROPENIA_GRAVE"         # <500
    NO_EVALUABLE = "RAN_NO_EVALUABLE"


class NivelEosinofilia(str, Enum):
    NORMAL = "EOSINOFILIA_NORMAL"           # <400
    LEVE = "EOSINOFILIA_LEVE"               # 400-1500
    MODERADA = "EOSINOFILIA_MODERADA"       # 1501-3000
    GRAVE = "EOSINOFILIA_GRAVE"             # >3000
    NO_EVALUABLE = "RAE_NO_EVALUABLE"


# Umbrales RAN (cel/mm3) — Norma MINSAL 2018, Sección 4
RAN_MINIMO_INICIO = 1500
RAN_ALARMA_1 = 1000
RAN_ALARMA_2 = 500

# Umbrales RAE (cel/mm3) — Norma MINSAL 2018, Anexo 2
RAE_LEVE = 400
RAE_MODERADA = 1500
RAE_GRAVE = 3000

VIGENCIA_DIAS = {
    FaseNorma.SEMANAL: 5,
    FaseNorma.QUINCENAL: 5,   # no explicitado en la norma para quincenal; criterio conservador. VALIDAR institucionalmente.
    FaseNorma.MENSUAL: 7,
}

DOSIS_MAXIMA_MG_DIA = 900
UMBRAL_BIMESTRE_DIAS = 60  # regla local de plataforma, NO normativa


# ============================================================
# ENTIDADES
# ============================================================

@dataclass
class Hemograma:
    fecha: date
    ran: Optional[float] = None
    rae: Optional[float] = None
    leucocitos: Optional[float] = None
    pct_neutrofilos: Optional[float] = None
    pct_eosinofilos: Optional[float] = None

    def ran_efectivo(self) -> Optional[float]:
        if self.ran is not None:
            return self.ran
        if self.leucocitos is not None and self.pct_neutrofilos is not None:
            return round(self.leucocitos * self.pct_neutrofilos / 100)
        return None

    def rae_efectivo(self) -> Optional[float]:
        if self.rae is not None:
            return self.rae
        if self.leucocitos is not None and self.pct_eosinofilos is not None:
            return round(self.leucocitos * self.pct_eosinofilos / 100)
        return None


@dataclass
class RegistroPaciente:
    rut: str
    fecha_dispensacion: date
    dosis_mg_dia: Optional[float]
    fase_norma: Optional[FaseNorma]
    paciente_nuevo: bool
    en_plataforma_minsal: bool
    ingreso_bimensual: bool
    fecha_ultimo_ingreso_plataforma: Optional[date]
    hemograma: Optional[Hemograma]
    neutropenia_etnica_benigna: bool = False
    dias_suspension_previa: Optional[int] = None


@dataclass
class Dictamen:
    rut: str
    accion_plataforma: AccionPlataforma
    nivel_alarma_ran: NivelAlarmaRAN
    nivel_eosinofilia: NivelEosinofilia
    ran_efectivo: Optional[float]
    rae_efectivo: Optional[float]
    hemograma_vigente: Optional[bool]
    observaciones: list[str] = field(default_factory=list)


# ============================================================
# CAPA 2 — VALIDACIONES CLÍNICO-NORMATIVAS
# ============================================================

def clasificar_ran(ran: Optional[float], neb: bool) -> tuple[NivelAlarmaRAN, list[str]]:
    obs = []
    if ran is None:
        return NivelAlarmaRAN.NO_EVALUABLE, ["RAN no disponible ni calculable (falta leucocitos y/o %neutrófilos). Solicitar dato al laboratorio."]
    if ran >= RAN_MINIMO_INICIO:
        return NivelAlarmaRAN.SIN_ALARMA, obs
    if ran >= RAN_ALARMA_1:
        if neb:
            obs.append(f"RAN {ran:.0f}: bajo 1.500 pero paciente con neutropenia étnica benigna documentada (excepción normativa, requiere plan con hematología).")
            return NivelAlarmaRAN.SIN_ALARMA, obs
        obs.append(f"RAN {ran:.0f}: ALARMA 1 (neutropenia leve, 1.000-1.499). Norma: hemograma c/48h hasta RAN>1.500, mantener tratamiento. Notificar a equipo tratante y Unidad de Mitigación de Riesgos.")
        return NivelAlarmaRAN.ALARMA_1_LEVE, obs
    if ran >= RAN_ALARMA_2:
        obs.append(f"RAN {ran:.0f}: ALARMA 2 (neutropenia moderada, 500-999). Norma: SUSPENDER clozapina, hemograma diario hasta RAN>1.000. Consultar hematología. Notificación obligatoria a Unidad de Mitigación de Riesgos. ESCALAR DE INMEDIATO.")
        return NivelAlarmaRAN.ALARMA_2_MODERADA, obs
    obs.append(f"RAN {ran:.0f}: ALARMA 3 (neutropenia grave/agranulocitosis, <500). Norma: suspensión inmediata + criterios de hospitalización. EMERGENCIA — escalar de inmediato.")
    return NivelAlarmaRAN.ALARMA_3_GRAVE, obs


def clasificar_eosinofilos(rae: Optional[float]) -> tuple[NivelEosinofilia, list[str]]:
    """Umbrales y conducta según Norma Técnica MINSAL 2018, Anexo 2 (eosinofilia)."""
    if rae is None:
        return NivelEosinofilia.NO_EVALUABLE, []  # RAE es secundario: no se alerta por ausencia, solo se omite clasificación
    if rae < RAE_LEVE:
        return NivelEosinofilia.NORMAL, []
    if rae <= RAE_MODERADA:
        return NivelEosinofilia.LEVE, [f"RAE {rae:.0f}: eosinofilia leve (400-1.500). Norma: control según esquema habitual, observar."]
    if rae <= RAE_GRAVE:
        return NivelEosinofilia.MODERADA, [f"RAE {rae:.0f}: eosinofilia moderada (1.501-3.000). Norma: control semanal, evaluar ajuste de dosis, evaluación por medicina interna."]
    return NivelEosinofilia.GRAVE, [f"RAE {rae:.0f}: eosinofilia grave (>3.000). Norma: control c/4 días, SUSPENDER clozapina, evaluación medicina interna/hematología, informar a Unidad de Mitigación de Riesgos."]


def validar_vigencia(hemograma: Hemograma, fase: FaseNorma, fecha_ref: date) -> tuple[bool, list[str]]:
    antiguedad = (fecha_ref - hemograma.fecha).days
    limite = VIGENCIA_DIAS[fase]
    if antiguedad < 0:
        return False, [f"Fecha de hemograma posterior a la dispensación: revisar registro, posible error de digitación."]
    if antiguedad > limite:
        return False, [f"Hemograma con {antiguedad} días de antigüedad excede vigencia normativa de {limite} días para fase {fase.value}. Gestionar nueva toma de muestra."]
    return True, []


def validar_dosis(dosis: float) -> list[str]:
    if dosis > DOSIS_MAXIMA_MG_DIA:
        return [f"Dosis {dosis:.0f} mg/día excede máximo normativo de {DOSIS_MAXIMA_MG_DIA} mg/día. Verificar receta con médico tratante."]
    if dosis <= 0:
        return [f"Dosis inválida ({dosis}). Corregir registro."]
    return []


def validar_suspension(dias: Optional[int]) -> list[str]:
    if dias is None:
        return []
    if dias > 30:
        return [f"Suspensión previa de {dias} días (>30): Norma exige REINICIAR esquema de hemograma semanal por 18 semanas."]
    return [f"Suspensión previa de {dias} días (≤30): puede continuarse desde última dosis con hemograma semanal hasta alcanzarla."]


# ============================================================
# CAPA 1 + CONSOLIDACIÓN
# ============================================================

def _bimestre_vigente(reg: RegistroPaciente) -> Optional[int]:
    if reg.fecha_ultimo_ingreso_plataforma is None:
        return None
    dias = (reg.fecha_dispensacion - reg.fecha_ultimo_ingreso_plataforma).days
    return dias if 0 <= dias < UMBRAL_BIMESTRE_DIAS else None


def evaluar(reg: RegistroPaciente) -> Dictamen:
    obs: list[str] = []
    nivel_ran = NivelAlarmaRAN.NO_EVALUABLE
    nivel_eos = NivelEosinofilia.NO_EVALUABLE
    ran = None
    rae = None
    vigente = None

    if reg.hemograma is not None:
        ran = reg.hemograma.ran_efectivo()
        rae = reg.hemograma.rae_efectivo()
        nivel_ran, obs_ran = clasificar_ran(ran, reg.neutropenia_etnica_benigna)
        obs.extend(obs_ran)
        nivel_eos, obs_eos = clasificar_eosinofilos(rae)
        obs.extend(obs_eos)
        if reg.fase_norma is not None:
            vigente, obs_vig = validar_vigencia(reg.hemograma, reg.fase_norma, reg.fecha_dispensacion)
            obs.extend(obs_vig)

    if reg.dosis_mg_dia is not None:
        obs.extend(validar_dosis(reg.dosis_mg_dia))
    obs.extend(validar_suspension(reg.dias_suspension_previa))

    if reg.paciente_nuevo or not reg.en_plataforma_minsal:
        obs.append("Paciente nuevo o no registrado en plataforma MINSAL. Avisar para gestionar alta antes de ingresar hemograma.")
        return Dictamen(reg.rut, AccionPlataforma.ALERTAR_NUEVO, nivel_ran, nivel_eos, ran, rae, vigente, obs)

    if reg.dosis_mg_dia is None or reg.fase_norma is None:
        faltante = [n for n, v in (("dosis", reg.dosis_mg_dia), ("fase de tratamiento", reg.fase_norma)) if v is None]
        obs.append(f"Dato(s) incompleto(s): {', '.join(faltante)}. No ingresar hemograma; feedback a equipo tratante.")
        return Dictamen(reg.rut, AccionPlataforma.NO_INGRESAR_DATOS_INCOMPLETOS, nivel_ran, nivel_eos, ran, rae, vigente, obs)

    dias_bim = _bimestre_vigente(reg) if reg.ingreso_bimensual else None
    if dias_bim is not None:
        obs.append(f"Ingreso bimensual vigente (último hace {dias_bim} días; regla local, no normativa). No insistir en reingreso; feedback informativo.")
        return Dictamen(reg.rut, AccionPlataforma.NO_INSISTIR_BIMENSUAL, nivel_ran, nivel_eos, ran, rae, vigente, obs)

    if reg.hemograma is None:
        obs.append("Sin hemograma en pestaña Laboratorio HCE SSASur. Avisar para gestión de toma de muestra.")
        return Dictamen(reg.rut, AccionPlataforma.ALERTAR_SIN_HEMOGRAMA, nivel_ran, nivel_eos, ran, rae, vigente, obs)

    ran_txt = f"RAN {ran:.0f}" if ran is not None else "RAN no calculable"
    rae_txt = f"RAE {rae:.0f}" if rae is not None else "RAE no calculable"
    obs.append(f"Hemograma {reg.hemograma.fecha.isoformat()} disponible ({ran_txt}, {rae_txt}), datos completos. Ingresar a plataforma MINSAL.")
    return Dictamen(reg.rut, AccionPlataforma.INGRESAR, nivel_ran, nivel_eos, ran, rae, vigente, obs)
