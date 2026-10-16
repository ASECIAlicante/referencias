"""Validacion estructural del JS inline de index.html sin depender de node.

Tokeniza saltandose strings, templates, comentarios y literales regex, y
comprueba el balance de (), [] y {}. No es un parser completo, pero detecta
exactamente el tipo de dano que puede causar una edicion a mano del bloque JS
embebido en build_html.py: una llave o un parentesis descolgados.

Comprueba tambien que cada funcion declarada se referencie en algun otro sitio
(deteccion de funciones muertas) y que no haya declaraciones duplicadas.
"""
import re
import sys
from collections import Counter

_html = open('index.html', encoding='utf-8').read()
_blocks = re.findall(r'<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>', _html, re.S)
if len(_blocks) != 1:
    print('se esperaba 1 bloque <script> inline en index.html, encontrados %d' % len(_blocks))
    sys.exit(1)
src = _blocks[0]

PAIRS = {')': '(', ']': '[', '}': '{'}
stack = []
i = 0
n = len(src)
line = 1
errors = []

# Un '/' abre regex solo si el token significativo anterior no puede cerrar una
# expresion. Sin esta distincion, una division como (a)/b se leeria como regex.
prev_significant = ''

while i < n:
    c = src[i]
    if c == '\n':
        line += 1
        i += 1
        continue
    # Comentarios
    if c == '/' and i + 1 < n and src[i + 1] == '/':
        j = src.find('\n', i)
        i = n if j == -1 else j
        continue
    if c == '/' and i + 1 < n and src[i + 1] == '*':
        j = src.find('*/', i + 2)
        if j == -1:
            errors.append('linea %d: comentario /* sin cerrar' % line)
            break
        line += src.count('\n', i, j)
        i = j + 2
        continue
    # Strings y templates
    if c in '"\'`':
        quote = c
        j = i + 1
        while j < n:
            if src[j] == '\\':
                j += 2
                continue
            if src[j] == quote:
                break
            if src[j] == '\n':
                if quote != '`':
                    errors.append('linea %d: string sin cerrar' % line)
                line += 1
            j += 1
        if j >= n:
            errors.append('linea %d: string sin cerrar al final del fichero' % line)
            break
        i = j + 1
        prev_significant = 'x'
        continue
    # Literal regex
    if c == '/' and prev_significant not in ('x', ')', ']', '}'):
        j = i + 1
        in_class = False
        ok = False
        while j < n and src[j] != '\n':
            if src[j] == '\\':
                j += 2
                continue
            if src[j] == '[':
                in_class = True
            elif src[j] == ']':
                in_class = False
            elif src[j] == '/' and not in_class:
                ok = True
                break
            j += 1
        if ok:
            i = j + 1
            prev_significant = 'x'
            continue
    if c in '([{':
        stack.append((c, line))
        prev_significant = c
        i += 1
        continue
    if c in ')]}':
        if not stack:
            errors.append('linea %d: cierre %r sin apertura' % (line, c))
        else:
            op, opline = stack.pop()
            if op != PAIRS[c]:
                errors.append('linea %d: cierre %r no corresponde a %r abierto en linea %d'
                              % (line, c, op, opline))
        prev_significant = c
        i += 1
        continue
    if not c.isspace():
        prev_significant = 'x'
    i += 1

for op, opline in stack:
    errors.append('linea %d: %r abierto y nunca cerrado' % (opline, op))

# --- Declaraciones duplicadas y funciones sin usar ---
decls = Counter(re.findall(r'\bfunction\s+([A-Za-z_$][\w$]*)\s*\(', src))
dups = [k for k, v in decls.items() if v > 1]
unused = []
for name in decls:
    # Referencias = apariciones del identificador menos la propia declaracion.
    refs = len(re.findall(r'\b%s\b' % re.escape(name), src)) - decls[name]
    if refs == 0:
        unused.append(name)

print('bloque JS: %d caracteres, %d lineas, %d funciones declaradas'
      % (len(src), src.count('\n') + 1, len(decls)))
if dups:
    print('AVISO funciones declaradas mas de una vez: %s' % ', '.join(sorted(dups)))
if unused:
    print('AVISO funciones declaradas y nunca referenciadas: %s' % ', '.join(sorted(unused)))
if errors:
    print('\nERRORES DE BALANCE:')
    for e in errors[:40]:
        print('  ' + e)
    sys.exit(1)
print('BALANCE OK: (), [] y {} cuadran en todo el bloque')
