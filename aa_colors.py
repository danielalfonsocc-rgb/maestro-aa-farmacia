"""Paleta de colores compartida para los niveles de Criticidad de AA.

Fuente unica de verdad usada por maestro_aa.py y app_pedidos.py, para que
todos los Excel generados (Consolidado, Resumen, descargas del dashboard)
muestren siempre el mismo color para el mismo nivel de criticidad.

Escalas de Criticidad usadas en el proyecto:
  Crit_Farm (crit_farm): 5-OK, 1-CRITICO, 2-URGENTE, 3-MODERADO, 4-BAJO
  Crit_Bod  (crit_bod) : 5-OK, 1-CRITICO, 2-URGENTE, 3-ALTO, 4-MODERADO, 5-BAJO

Los pares 3-MODERADO/4-MODERADO y 4-BAJO/5-BAJO comparten color: ocupan la
misma posicion relativa (penultimo / ultimo escalon de urgencia) en cada
escala, aunque tengan distinto numero por tener Bodega un escalon mas.
"""

from openpyxl.styles import PatternFill, Font

CRIT_FILL_HEX = {
    '1-CRITICO' : 'B71C1C',  # rojo fuerte   - sin stock / sin respaldo
    '2-URGENTE' : 'FFCDD2',  # rojo claro
    '3-ALTO'    : 'FFE0B2',  # naranja medio  (Bodega: cobertura < 1 semana laboral)
    '3-MODERADO': 'FFF3E0',  # naranja claro  (Farmacia)
    '4-MODERADO': 'FFF3E0',  # naranja claro  (Bodega: cobertura < 1 ciclo de 2 semanas)
    '4-BAJO'    : 'FFF9C4',  # amarillo       (Farmacia)
    '5-BAJO'    : 'FFF9C4',  # amarillo       (Bodega: cobertura >= 1 ciclo de 2 semanas)
    '5-OK'      : 'E8F5E9',  # verde          - sin necesidad
}

FONT_CRITICO = Font(bold=True, color='FFFFFF', name='Arial', size=11)


def crit_fill(crit, default='FFFFFF'):
    """PatternFill solido para un nivel de Criticidad (1-CRITICO, 2-URGENTE, ...)."""
    return PatternFill('solid', fgColor=CRIT_FILL_HEX.get(str(crit), default))


# ─────────────────────────────────────────────
# NIVEL DE CRITICIDAD — normaliza las DOS escalas del proyecto a 1..5
# ─────────────────────────────────────────────
# El proyecto usa dos formas de escribir la Criticidad:
#   Pedidos   : '1-CRITICO', '2-URGENTE', '3-ALTO'/'3-MODERADO', '4-...', '5-...'
#   Faltantes : '[CRITICO] CRITICO - SIN RESPALDO', '[ALTO] ...',
#               '[MODERADO] ...', '[BAJO] ...'
# crit_nivel() las reduce a un solo numero (1 = mas urgente, 5 = sin necesidad)
# para que la app y los Excel coloreen y ordenen igual sin importar la escala.
_NIVEL_HEX = {1: 'B71C1C', 2: 'FFCDD2', 3: 'FFE0B2', 4: 'FFF9C4', 5: 'E8F5E9'}


def crit_nivel(crit) -> int:
    """Devuelve 1..5 para cualquier etiqueta de Criticidad (ambas escalas)."""
    c = str(crit).strip().upper()
    # Escala numerica de pedidos: el prefijo "N-" manda
    if len(c) >= 2 and c[0] in '12345' and c[1] == '-':
        return int(c[0])
    # Escala de faltantes (texto): CRITICO > ALTO > MODERADO > BAJO
    if 'CRITICO'  in c: return 1
    if 'URGENTE'  in c: return 2
    if 'ALTO'     in c: return 2
    if 'MODERADO' in c: return 3
    if 'BAJO'     in c: return 4
    return 5


def crit_hex(crit, default='F8F8F8'):
    """Color hex representativo para cualquier Criticidad (via crit_nivel)."""
    return _NIVEL_HEX.get(crit_nivel(crit), default)


def is_critico(crit):
    return crit_nivel(crit) == 1
