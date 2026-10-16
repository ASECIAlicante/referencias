"""Comprobaciones de integridad de los artefactos generados (style.css + index.html).

1. Balance de llaves en style.css: una regla descolgada por una edicion a mano
   del bloque CSS de build_html.py no rompe el build, solo deja la hoja en
   silencio a medias, asi que conviene detectarlo aqui.
2. Toda clase usada en el markup de index.html que no tenga regla en style.css
   ni se inyecte desde JS (informativo, no error).
3. Todo getElementById / querySelector('#id') del JS tiene su id en el markup.
"""
import re
import sys

css = open('style.css', encoding='utf-8').read()
html = open('index.html', encoding='utf-8').read()

errors = []

# --- 1) Balance de llaves en CSS (ignorando comentarios) ---
css_nc = re.sub(r'/\*.*?\*/', '', css, flags=re.S)
depth = 0
for ch in css_nc:
    if ch == '{':
        depth += 1
    elif ch == '}':
        depth -= 1
        if depth < 0:
            errors.append('style.css: "}" de mas')
            break
if depth > 0:
    errors.append('style.css: %d bloque(s) "{" sin cerrar' % depth)

# --- 2) ids referenciados desde el JS que no existen en el markup ---
script = re.search(r'<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>', html, re.S)
js = script.group(1) if script else ''
markup = html.replace(js, '')
declared = set(re.findall(r'\bid="([^"]+)"', markup))
referenced = set(re.findall(r"getElementById\('([^']+)'\)", js))
referenced |= set(re.findall(r"querySelector\('#([A-Za-z][\w-]*)'\)", js))
missing = sorted(referenced - declared)
if missing:
    errors.append('ids referenciados por el JS que no existen en el markup: %s'
                  % ', '.join(missing))

# --- 3) ids declarados que nadie usa (informativo) ---
# Los ids generados por datos (paneles de hoja, secciones, tarjetas, filas de
# LOB) se resuelven en caliente con getElementById(slug), donde slug es una
# variable: un grep estatico no los ve y los daria todos por huerfanos. Se
# excluyen por patron para que la lista que queda sea accionable.
DYNAMIC_ID_RE = re.compile(r'^(sec-\d+-\d+(-card)?|lobrow-[\w-]+|home)$')
css_ids = set(re.findall(r'#([A-Za-z][\w-]*)', css))
href_ids = set(re.findall(r'(?:href|aria-labelledby|xlink:href)="#?([\w-]+)"', markup))
url_ids = set(re.findall(r'url\(#([\w-]+)\)', markup))
# Los slugs de hoja son tambien destinos de navegacion: aparecen en data-target
# / data-lob de los botones del menu.
nav_targets = set(re.findall(r'data-(?:target|lob)="([\w-]+)"', markup))
orphan_ids = sorted(i for i in declared - referenced - css_ids - href_ids - url_ids - nav_targets
                    if not DYNAMIC_ID_RE.match(i))

print('style.css: %d caracteres, llaves balanceadas' % len(css))
print('ids en markup: %d | referenciados desde JS: %d' % (len(declared), len(referenced)))
if orphan_ids:
    print('ids sin ninguna referencia (JS, CSS, href, aria, url()): %s' % ', '.join(orphan_ids))
else:
    print('ids sin referencia: ninguno')

if errors:
    print('\nERRORES:')
    for e in errors:
        print('  ' + e)
    sys.exit(1)
print('INTEGRIDAD OK')
