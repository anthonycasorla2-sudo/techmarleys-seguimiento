"""Imprime cuántos segundos faltan para la próxima revisión encadenada (HH:07
dentro del horario de rastreador.py), o -1 si falta más de ESPERA_MAXIMA.

Lo usa el workflow de GitHub para encadenar una revisión con la siguiente.
No necesita librerías instaladas.
"""

from datetime import datetime, timedelta

import rastreador

MINUTO_REVISION = 7
ESPERA_MAXIMA = 3 * 3600  # de noche no se espera: la mañana la arranca el horario de GitHub


def segundos_hasta_proxima(ahora=None):
    ahora = ahora or datetime.now()
    for dias in (0, 1):
        for hora in range(rastreador.HORA_INICIO, rastreador.HORA_FIN + 1):
            t = (ahora + timedelta(days=dias)).replace(hour=hora, minute=MINUTO_REVISION, second=0, microsecond=0)
            if t > ahora + timedelta(minutes=5):
                espera = int((t - ahora).total_seconds())
                return espera if espera <= ESPERA_MAXIMA else -1
    return -1


if __name__ == "__main__":
    print(segundos_hasta_proxima())
