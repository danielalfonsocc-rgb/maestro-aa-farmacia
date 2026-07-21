"""Suite de pruebas V3 — incorpora clasificación RAE (eosinofilia)."""
from datetime import date, timedelta
from motor_reglas_clozapina_v3 import (
    RegistroPaciente, Hemograma, FaseNorma, evaluar,
    AccionPlataforma, NivelAlarmaRAN, NivelEosinofilia
)

hoy = date(2026, 7, 20)

def hg(dias_atras, ran=None, rae=None, leu=None, pct_n=None, pct_e=None):
    return Hemograma(fecha=hoy - timedelta(days=dias_atras), ran=ran, rae=rae,
                      leucocitos=leu, pct_neutrofilos=pct_n, pct_eosinofilos=pct_e)

casos = [
    ("T01 caso base, RAN y RAE normales", RegistroPaciente(
        "44.444.444-4", hoy, 300, FaseNorma.MENSUAL, False, True, False,
        hoy - timedelta(days=31), hg(3, ran=3200, rae=200)),
        AccionPlataforma.INGRESAR, NivelAlarmaRAN.SIN_ALARMA, NivelEosinofilia.NORMAL),

    ("T02 paciente nuevo con ALARMA 2 RAN", RegistroPaciente(
        "11.111.111-1", hoy, 100, FaseNorma.SEMANAL, True, False, False,
        None, hg(1, ran=750, rae=150)),
        AccionPlataforma.ALERTAR_NUEVO, NivelAlarmaRAN.ALARMA_2_MODERADA, NivelEosinofilia.NORMAL),

    ("T03 fase incompleta", RegistroPaciente(
        "22.222.222-2", hoy, 300, None, False, True, False,
        hoy - timedelta(days=30), hg(2, ran=2500, rae=100)),
        AccionPlataforma.NO_INGRESAR_DATOS_INCOMPLETOS, NivelAlarmaRAN.SIN_ALARMA, NivelEosinofilia.NORMAL),

    ("T04 bimensual vigente", RegistroPaciente(
        "33.333.333-3", hoy, 400, FaseNorma.MENSUAL, False, True, True,
        hoy - timedelta(days=20), hg(3, ran=2800, rae=300)),
        AccionPlataforma.NO_INSISTIR_BIMENSUAL, NivelAlarmaRAN.SIN_ALARMA, NivelEosinofilia.NORMAL),

    ("T05 sin hemograma en HCE", RegistroPaciente(
        "55.555.555-5", hoy, 300, FaseNorma.MENSUAL, False, True, False,
        hoy - timedelta(days=35), None),
        AccionPlataforma.ALERTAR_SIN_HEMOGRAMA, NivelAlarmaRAN.NO_EVALUABLE, NivelEosinofilia.NO_EVALUABLE),

    ("T06 RAN y RAE calculados desde leucocitos x %", RegistroPaciente(
        "66.666.666-6", hoy, 300, FaseNorma.MENSUAL, False, True, False,
        hoy - timedelta(days=32), hg(4, leu=6000, pct_n=55, pct_e=8)),
        AccionPlataforma.INGRESAR, NivelAlarmaRAN.SIN_ALARMA, NivelEosinofilia.LEVE),  # RAN=3300, RAE=480

    ("T07 eosinofilia MODERADA (RAE 2000)", RegistroPaciente(
        "77.777.777-7", hoy, 300, FaseNorma.MENSUAL, False, True, False,
        hoy - timedelta(days=30), hg(1, ran=3000, rae=2000)),
        AccionPlataforma.INGRESAR, NivelAlarmaRAN.SIN_ALARMA, NivelEosinofilia.MODERADA),

    ("T08 eosinofilia GRAVE (RAE 3500) + RAN normal", RegistroPaciente(
        "88.888.888-8", hoy, 300, FaseNorma.MENSUAL, False, True, False,
        hoy - timedelta(days=30), hg(1, ran=2900, rae=3500)),
        AccionPlataforma.INGRESAR, NivelAlarmaRAN.SIN_ALARMA, NivelEosinofilia.GRAVE),

    ("T09 RAE ausente no bloquea clasificación RAN", RegistroPaciente(
        "99.999.999-9", hoy, 300, FaseNorma.MENSUAL, False, True, False,
        hoy - timedelta(days=31), hg(2, ran=1200)),  # sin rae
        AccionPlataforma.INGRESAR, NivelAlarmaRAN.ALARMA_1_LEVE, NivelEosinofilia.NO_EVALUABLE),

    ("T10 doble alarma: RAN alarma 1 + eosinofilia grave simultáneas", RegistroPaciente(
        "10.101.010-1", hoy, 300, FaseNorma.MENSUAL, False, True, False,
        hoy - timedelta(days=30), hg(1, ran=1100, rae=3200)),
        AccionPlataforma.INGRESAR, NivelAlarmaRAN.ALARMA_1_LEVE, NivelEosinofilia.GRAVE),
]

fallos = 0
for nombre, reg, acc_esp, ran_esp, eos_esp in casos:
    d = evaluar(reg)
    ok = (d.accion_plataforma == acc_esp and d.nivel_alarma_ran == ran_esp and d.nivel_eosinofilia == eos_esp)
    if not ok:
        fallos += 1
    print(f"[{'PASS' if ok else 'FAIL'}] {nombre}")
    print(f"       accion={d.accion_plataforma.value} | RAN_alarma={d.nivel_alarma_ran.value} | RAE_alarma={d.nivel_eosinofilia.value}")
    print(f"       RAN={d.ran_efectivo} | RAE={d.rae_efectivo} | vigente={d.hemograma_vigente}")
    for o in d.observaciones:
        print(f"       - {o[:130]}")
    print()

print(f"{'='*80}\nResultado: {len(casos)-fallos}/{len(casos)} PASS")
