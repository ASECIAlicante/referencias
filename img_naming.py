"""Reglas de nomenclatura de imagenes compartidas entre build_html.py y
resize_inplace.py.

Los dos scripts necesitan reconocer el mismo patron de nombre duplicado que
macOS/iCloud/los navegadores generan al descargar dos veces el mismo fichero
("<base> 2.webp", "<base> copia.webp", "<base> copy.webp"...). Antes el mismo
patron vivia copiado letra por letra en ambos ficheros, con riesgo de que un
ajuste futuro (p. ej. para cubrir un sufijo nuevo) se aplicara en uno y se
olvidara en el otro, dejando a build_html.py y resize_inplace.py en desacuerdo
sobre que es un duplicado.
"""
import re

# Siempre llevan un espacio antes del sufijo; un nombre de imagen bueno (solo
# el EAN/referencia) nunca lo lleva, asi que el espacio es la senal
# inequivoca de duplicado y no puede colisionar con una imagen valida.
DUP_STEM_RE = re.compile(
    r'^(?P<base>.+?)[ ]+(?:\d+|cop(?:ia|y)(?:[ ]+\d+)?)$',
    re.IGNORECASE,
)
