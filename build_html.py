# -*- coding: utf-8 -*-
"""
Generador automático del catálogo HTML de referencias Apple a partir de Database.xlsx.
Estructura de navegación por niveles (L1 categorías, L2 hojas, L3 modelos).
Incluye generador e interactividad de códigos de barras EAN/UPC vía SVG.
"""

import openpyxl, json, colorsys, re, os, html, datetime
from img_naming import DUP_STEM_RE

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'Database.xlsx')
HTML_PATH = os.path.join(BASE_DIR, 'index.html')

# Diccionario de busqueda ES->EN. Se lee en CADA build desde el fichero
# "Colors" del proyecto (.rtf, .txt o sin extension), de modo que basta con
# editar ese fichero y relanzar el build para que el buscador se actualice.
# Formato de cada linea:  <termino en ingles> = <termino en espanol>
# Las lineas SIN "=" se ignoran (aun sin traducir), igual que las que tengan
# cualquiera de los dos lados vacio. Un mismo termino ingles puede repetirse
# con varias traducciones (Pencil = Lapiz / Pencil = Boli) y una misma
# traduccion puede apuntar a varios terminos ingleses (Blanco estrella ->
# Starlight y Star White): ambos casos se acumulan.
COLORS_CANDIDATES = ('Colors.rtf', 'Colors.txt', 'Colors', 'Colours.rtf', 'Colours.txt')

_ACCENTS = str.maketrans({
    'á': 'a', 'à': 'a', 'ä': 'a', 'â': 'a', 'ã': 'a',
    'é': 'e', 'è': 'e', 'ë': 'e', 'ê': 'e',
    'í': 'i', 'ì': 'i', 'ï': 'i', 'î': 'i',
    'ó': 'o', 'ò': 'o', 'ö': 'o', 'ô': 'o', 'õ': 'o',
    'ú': 'u', 'ù': 'u', 'ü': 'u', 'û': 'u',
    'ñ': 'n', 'ç': 'c',
})

_ALPHA_WORD = re.compile(r'^[a-z]+$')


def strip_accents(s):
    """Minusculas y sin acentos, SIN tocar plural ni genero. Es el nivel que se
    usa para casar texto literal (P/N, EAN, ECI, descripcion) y los terminos
    del diccionario contra la fila: conserva la grafia real del dato."""
    return str(s).lower().translate(_ACCENTS)


def fold_es(s):
    """strip_accents() y ademas sin marca de plural ni de genero. Debe producir
    EXACTAMENTE el mismo resultado que foldES() en el JavaScript del buscador:
    si las dos divergen, el diccionario deja de casar.
    'Pequeña' -> 'pequeno', 'Burdeos' -> 'burdeo', 'Rosa nube' -> 'roso nube'.

    Solo se recorta en palabras puramente alfabeticas. Un P/N ('MHFA4Y/A'), un
    EAN, una referencia ECI o una talla ('S/M') llevan digitos o barras y se
    dejan intactos. Este nivel se usa UNICAMENTE para buscar la clave en el
    diccionario, nunca para casar contra la fila.
    """
    out = []
    for w in strip_accents(s).split():
        if _ALPHA_WORD.match(w):
            if len(w) > 3 and w.endswith('s'):
                w = w[:-1]
            if len(w) > 3 and w.endswith('a'):
                w = w[:-1] + 'o'
        out.append(w)
    return ' '.join(out)


def _rtf_to_text(raw):
    """Extrae el texto plano de un RTF simple (el que genera TextEdit)."""
    for grp in ('fonttbl', 'colortbl', 'expandedcolortbl', 'stylesheet', 'info'):
        raw = re.sub(r'\{\\\*?\\' + grp + r'[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', '', raw)
    # Caracteres acentuados escapados como \'e9 (cp1252)
    raw = re.sub(r"\\'([0-9a-fA-F]{2})",
                 lambda m: bytes([int(m.group(1), 16)]).decode('cp1252', 'replace'), raw)
    raw = raw.replace('\\\r\n', '\n').replace('\\\n', '\n')
    raw = re.sub(r'\\par[d]?\b', '\n', raw)
    raw = re.sub(r'\\[a-zA-Z]+-?\d*\s?', '', raw)
    return raw.replace('{', '').replace('}', '')


def load_term_synonyms():
    """Lee el fichero Colors y devuelve (dict, max_palabras, informe).

    Cada linea "A = B" declara que A y B son el MISMO concepto. No importa que
    lado es ingles y cual espanol, ni el orden de las lineas: los terminos se
    agrupan en clases de equivalencia, de modo que teclear cualquiera de ellos
    encuentra las filas escritas con cualquier otro. Las lineas que comparten
    un termino se fusionan solas ('Display = Pantalla' + 'Display = Monitor'
    -> pantalla, monitor y display son intercambiables).

    dict: {clave_normalizada: [termino, ...]} con todos los equivalentes.
    """
    path = next((p for p in (os.path.join(BASE_DIR, n) for n in COLORS_CANDIDATES)
                 if os.path.exists(p)), None)
    if not path:
        return {}, 1, 'Diccionario de busqueda: no se encontro el fichero Colors (buscador sin traduccion)'

    with open(path, 'rb') as f:
        data = f.read()
    text = data.decode('utf-8', 'replace')
    if text.lstrip().startswith('{\\rtf') or os.path.basename(path).lower().endswith('.rtf'):
        text = _rtf_to_text(data.decode('cp1252', 'replace'))

    # --- Union-find sobre los terminos normalizados ---
    parent, label = {}, {}

    def find(x):
        parent.setdefault(x, x)
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    skipped, lines_ok = 0, 0
    for line in text.splitlines():
        line = line.strip()
        if not line or '=' not in line:
            if line:
                skipped += 1
            continue
        left, _, right = line.partition('=')
        left, right = left.strip(), right.strip()
        if not left or not right:
            skipped += 1
            continue
        kl, kr = fold_es(left), fold_es(right)
        label.setdefault(kl, left.lower())
        label.setdefault(kr, right.lower())
        if kl != kr:
            union(kl, kr)
        lines_ok += 1

    classes = {}
    for term in list(parent):
        classes.setdefault(find(term), []).append(term)

    # Una clase de un solo termino no aporta nada: ese texto ya se encuentra
    # por coincidencia literal.
    syn = {}
    for members in classes.values():
        if len(members) < 2:
            continue
        vals = sorted({label.get(m, m) for m in members})
        for m in members:
            syn[m] = vals

    # Los terminos de varias palabras ("Blanco estrella") se indexan tambien
    # palabra a palabra, para que buscar solo "estrella" siga encontrando
    # Starlight. Esto se hace incluso si esa palabra YA es una entrada propia
    # ("Verde" = Green): se AMPLIA su lista de equivalentes (nunca se pisa,
    # solo se añade), para que buscar "verde" tambien encuentre "Verde oliva",
    # "Verde salvia" o "Verde azulado", y no solo las filas literalmente
    # etiquetadas "Verde"/"Green".
    for key, vals in list(syn.items()):
        words = key.split()
        if len(words) < 2:
            continue
        for w in words:
            if len(w) < 4:
                continue
            bucket = syn.setdefault(w, [])
            for v in vals:
                if v not in bucket:
                    bucket.append(v)

    # Plurales en consonante ("azul"->"azules", "lapiz"->"lapices"): al teclear
    # el plural, fold_es solo quita una 's' ("azules"->"azule", "lapices"->
    # "lapice"), forma que NO coincide con la clave singular. No se puede tocar
    # el plegado para arreglarlo sin romper los plurales en vocal (que ya
    # convergen: "verdes"->"verde", "estrellas"->"estrella"), porque "Xes" es
    # estructuralmente identico y su singular difiere. En su lugar registramos
    # esa forma plegada del plural como alias de la misma clase de equivalencia.
    _VOWELS = set('aeiou')
    for key, vals in list(syn.items()):
        if ' ' in key or len(key) <= 3 or not _ALPHA_WORD.match(key):
            continue
        last = key[-1]
        if last == 'z':
            alias = key[:-1] + 'ce'      # lapiz -> lapice (plegado de "lapices")
        elif last not in _VOWELS:
            alias = key + 'e'            # azul  -> azule  (plegado de "azules")
        else:
            continue                     # vocal: lo cubre ya el quitar una "s"
        if alias not in syn:
            syn[alias] = vals

    pairs = syn
    max_words = max((len(k.split()) for k in pairs), key=int, default=1)
    report = (f'Diccionario de busqueda: {len(pairs)} terminos en '
              f'{sum(1 for v in classes.values() if len(v) > 1)} grupos desde '
              f'{os.path.basename(path)} ({lines_ok} lineas leidas, '
              f'{skipped} sin "=" ignoradas)')
    return pairs, max_words, report


THEME = {
    0: 'FFFFFF', 1: '000000', 2: 'E8E8E8', 3: '0E2841',
    4: '156082', 5: 'E97132', 6: '196B24', 7: '0F9ED5',
    8: 'A02B93', 9: '4EA72E',
}

KNOWN_LABELS = {'p/n', 'eci', 'ean', 'wifi', 'lte', 'gps', 'cellular', 'gps + cellular', 'wi-fi'}
MONO_LABELS = {'p/n', 'eci', 'ean', 'wifi', 'lte', 'gps', 'cellular'}

# Marco de botones y tarjetas: el color de familia ya NO sale de una paleta
# fija, sino de los colores REALES de producto que trae cada celda de la hoja
# (el fondo de celda = color del producto). Se calcula en build, de modo que al
# actualizar Database.xlsx los marcos se recalculan solos. Un modelo con 4
# acabados (p.ej. Neo) genera un degradado de 4 colores; un MacBook Pro con 2
# acabados, de 2. Como los dos tamanos de un mismo modelo comparten acabados,
# su paleta coincide y la agrupacion por familia se mantiene de forma natural.
# Se limita el numero de paradas para que el degradado siga siendo legible.
FAM_GRAD_MAX_STOPS = 6
FAM_WASH_ALPHA = 0.14

# EAN cuyo relleno de color en la celda EAN no es un color de producto real,
# sino un aviso/marca comercial (p.ej. el cargador 30W en liquidacion resaltado
# en todas las hojas de accesorios) y que por tanto nunca debe sumarse a la
# paleta de colores de familia de ninguna tarjeta.
COLOR_EXCLUDED_EANS = {'0195949376511'}


_SIZE_TOKEN_RE = re.compile(r'\d+(?:[.,]\d+)?\s*(?:mm|"|”|ª)', re.IGNORECASE)
_CHIP_TOKEN_RE = re.compile(r'[AM]\d+(?:\s+(?:Pro|Max|Ultra))?')
_VARIANT_WORD_RE = re.compile(r'\b(Wifi|Wi-Fi|LTE|GPS|CELL|Cellular|XDR)\b', re.IGNORECASE)
# "Max"/"Plus" como escalon de TAMANO del mismo modelo: queremos que, p.ej.,
# "18 Pro Max" y "18 Pro" caigan en la misma familia. OJO: no se incluye "mini"
# porque en iPad "Mini 7" es un producto propio, no una variante de tamano de
# otro. El chip "M5 Max"/"M5 Ultra" de los Mac ya lo consume antes _CHIP_TOKEN_RE,
# asi que este paso solo afecta al "Max" suelto de los iPhone.
_SIZE_TIER_RE = re.compile(r'\b(?:Max|Plus)\b', re.IGNORECASE)

def family_key(title):
    """Reduce un titulo de modelo (p.ej. 'Air 13,6” M5', 'Pro 16” M5 Max',
    '42mm Titanium GPS + CELL', '18 Pro Max') a una clave de "familia"
    agrupando variantes del mismo producto base: se quitan tamanos (13,6”,
    42mm, 7ª), chips (M5, A18 Pro) que modifican a una palabra anterior,
    palabras de conectividad o acabado (Wifi/LTE/GPS/CELL/XDR) y el escalon de
    tamano Max/Plus. Un chip al PRINCIPIO del titulo (p.ej. 'A16 Wifi') se
    conserva, porque ahi el chip ES el nombre del producto, no un modificador
    (no hay 'Air'/'Pro' delante que haga de familia). Si tras limpiar no queda
    nada reconocible, se usa la primera palabra original como respaldo para no
    perder la agrupacion.
    """
    s = _SIZE_TOKEN_RE.sub(' ', title)

    def _chip_sub(m):
        return m.group(0) if m.start() == 0 else ''
    s = _CHIP_TOKEN_RE.sub(_chip_sub, s)

    s = _VARIANT_WORD_RE.sub(' ', s)
    s = _SIZE_TIER_RE.sub(' ', s)
    s = s.replace('+', ' ').replace('/', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    if not s:
        words = title.split()
        s = words[0] if words else title
    return s.lower()

def _hex_to_rgb(hexc):
    h = hexc.lstrip('#')
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

def _rgba(hexc, a):
    r, g, b = _hex_to_rgb(hexc)
    return f'rgba({r},{g},{b},{a})'

def _boost_color(hexc):
    """Sube la saturacion y acota la luminosidad a un rango medio-visible,
    para que el marco de color de la tarjeta/boton se perciba con contraste
    claro incluso cuando el color de producto real es un tono muy palido o
    casi blanco/negro (p.ej. plata, blanco estrella) — sobre todo notorio en
    el tema claro, donde esos tonos se confundian con el fondo del panel."""
    r, g, b = _hex_to_rgb(hexc)
    h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
    s = min(1.0, s * 1.8 + 0.18)
    l = min(0.62, max(0.3, l))
    r2, g2, b2 = colorsys.hls_to_rgb(h, l, s)
    return '#{:02X}{:02X}{:02X}'.format(
        round(max(0, min(255, r2 * 255))),
        round(max(0, min(255, g2 * 255))),
        round(max(0, min(255, b2 * 255))),
    )

def _rows_product_colors(rows, cols):
    """Colores de producto DISTINTOS de una lista de filas (de una sola
    subseccion), en orden de aparicion. Solo se mira el fondo de la celda EAN
    de cada fila: es la unica columna que identifica univocamente un
    producto. Otras celdas de la misma fila (P/N, ECI, notas...) pueden
    llevar relleno de color por motivos puramente informativos (p.ej. en
    Watch, para marcar disponibilidad) y no deben colarse en la paleta de
    familia. Es la fuente de verdad del color: viene de Database.xlsx, asi
    que se recalcula en cada build."""
    ean_idx = next((i for i, c in enumerate(cols) if c.strip().upper() == 'EAN'), None)
    if ean_idx is None:
        return []
    colors = []
    seen = set()
    for row in rows:
        col_pos = 0
        for c in row:
            span = c.get('span') or 1
            if col_pos == ean_idx:
                # Solo cuenta si la celda EAN tiene de verdad un EAN (no una
                # nota combinada u otra fila sin valor en esa columna que por
                # casualidad caiga en esa posicion con un relleno de color),
                # y si no es uno de los EAN excluidos explicitamente (p.ej. el
                # cargador en liquidacion, cuyo relleno de color es un aviso
                # de oferta y no un color de producto real).
                ean_val = c.get('v')
                if ean_val and ean_val.strip() and ean_val.strip() not in COLOR_EXCLUDED_EANS:
                    bg = c.get('bg')
                    if bg:
                        key = bg.upper()
                        if key not in seen:
                            seen.add(key)
                            colors.append(bg)
                break
            col_pos += span
    return colors

def section_product_colors(sec):
    """Colores de producto de una seccion/tarjeta para la paleta de familia.

    Una tarjeta puede llevar, ademas del propio producto, una o varias
    subsecciones de "Accesorios recomendados" (la banda amarilla del Excel).
    Esas filas de accesorios a veces llevan su propio relleno de color (p.ej.
    un resaltado verde de "recomendado" o naranja de oferta en el cargador),
    que no es un color de producto real y nunca debe sumarse al degradado de
    familia. El corte se hace a nivel de fila (ver 'color_rows' en
    process_normal_sheet): en cuanto aparece la fila "Accesorios
    recomendados" -- esa fila incluida -- se deja de mirar todo lo que venga
    despues en esa misma subseccion/tarjeta, tenga o no EAN con relleno."""
    groups = sec.get('groups')
    if groups:
        colors = []
        seen = set()
        for grp in groups:
            if grp.get('is_accessory'):
                break
            for c in _rows_product_colors(grp.get('color_rows', grp['rows']), grp['columns']):
                key = c.upper()
                if key not in seen:
                    seen.add(key)
                    colors.append(c)
        return colors
    if sec.get('is_accessory'):
        return []
    return _rows_product_colors(sec.get('color_rows', sec['rows']), sec.get('columns') or [])

def _union_colors(list_of_color_lists):
    out = []
    seen = set()
    for cols in list_of_color_lists:
        for col in cols:
            key = col.upper()
            if key in seen:
                continue
            seen.add(key)
            out.append(col)
    return out

def _sample_colors(colors, cap=FAM_GRAD_MAX_STOPS):
    """Si una familia tiene demasiados colores (p.ej. hojas de correas), se
    muestrean de forma uniforme para que el degradado no se vuelva un arcoiris
    ilegible, conservando el primero y el ultimo."""
    if len(colors) <= cap:
        return colors
    step = (len(colors) - 1) / (cap - 1)
    return [colors[round(i * step)] for i in range(cap)]

def gradient_vars(colors):
    """A partir de la lista de colores de producto de una familia, devuelve
    (grad, wash) para las CSS custom props --fam-grad (paradas solidas del
    marco) y --fam-wash (las mismas paradas a baja opacidad, para el relleno
    sutil). Devuelve (None, None) si la familia no tiene ningun color: en ese
    caso el elemento se queda sin marco de familia."""
    if not colors:
        return None, None
    cols = [_boost_color(c) for c in _sample_colors(colors)]
    if len(cols) == 1:
        c = cols[0]
        return f'{c}, {c}', f'{_rgba(c, FAM_WASH_ALPHA)}, {_rgba(c, FAM_WASH_ALPHA)}'
    grad = ', '.join(cols)
    wash = ', '.join(_rgba(c, FAM_WASH_ALPHA) for c in cols)
    return grad, wash

def family_palettes_for(sections):
    """Dada una lista de secciones (en orden de aparicion), agrupa por
    family_key y devuelve (pos_palette, n_familias): pos_palette[i] es la union
    de colores de producto de la familia a la que pertenece la seccion i. Asi
    todas las variantes de una familia comparten exactamente la misma paleta
    (p.ej. MacBook Air 13 y 15, o iPad Air 11 y 13)."""
    members = {}
    order = []
    for i, sec in enumerate(sections):
        k = family_key(sec['title'])
        if k not in members:
            members[k] = []
            order.append(k)
        members[k].append(i)
    palette = {}
    for k in order:
        palette[k] = _union_colors([section_product_colors(sections[i]) for i in members[k]])
    pos_palette = {i: palette[family_key(sec['title'])] for i, sec in enumerate(sections)}
    return pos_palette, len(order)

def fam_style_attr(grad, wash):
    """style="--fam-grad:...;--fam-wash:..." o '' si no hay colores."""
    if not grad:
        return ''
    return f' style="--fam-grad:{grad};--fam-wash:{wash}"'

def apply_tint(hex_color, tint):
    if tint is None or tint == 0:
        return hex_color
    r = int(hex_color[0:2], 16) / 255.0
    g = int(hex_color[2:4], 16) / 255.0
    b = int(hex_color[4:6], 16) / 255.0
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if tint < 0:
        l = l * (1.0 + tint)
    else:
        l = l * (1.0 - tint) + tint
    r2, g2, b2 = colorsys.hls_to_rgb(h, l, s)
    return '{:02X}{:02X}{:02X}'.format(int(round(r2*255)), int(round(g2*255)), int(round(b2*255)))

def cell_color(cell):
    fill = cell.fill
    if fill is None or fill.fgColor is None:
        return None
    fg = fill.fgColor
    try:
        if fill.patternType is None:
            return None
    except Exception:
        pass
    if fg.type == 'rgb':
        rgb = fg.rgb
        if isinstance(rgb, str) and len(rgb) >= 6:
            hexv = rgb[-6:]
            if hexv.upper() == '000000' and rgb.upper() == '00000000':
                return None
            return '#' + hexv
        return None
    elif fg.type == 'theme':
        theme_idx = fg.theme
        base = THEME.get(theme_idx, 'FFFFFF')
        tint = fg.tint or 0.0
        return '#' + apply_tint(base, tint)
    return None

def luminance(hexcolor):
    hexcolor = hexcolor.lstrip('#')
    r = int(hexcolor[0:2], 16) / 255.0
    g = int(hexcolor[2:4], 16) / 255.0
    b = int(hexcolor[4:6], 16) / 255.0
    def lin(c):
        return c/12.92 if c <= 0.03928 else ((c+0.055)/1.055) ** 2.4
    r, g, b = lin(r), lin(g), lin(b)
    return 0.2126*r + 0.7152*g + 0.0722*b

def text_color_for(bg_hex):
    if not bg_hex:
        return None
    lum = luminance(bg_hex)
    return '#111318' if lum > 0.5 else '#FFFFFF'

def esc(s):
    return html.escape(str(s), quote=True)

def cell_to_dict(cell, with_bold=False):
    """Construye el dict {'v','bg','fg'[,'bold']} comun a hojas normales y
    matrices, a partir de una celda openpyxl."""
    val = cell.value
    bg = cell_color(cell)
    d = {
        'v': '' if val is None else str(val),
        'bg': bg,
        'fg': text_color_for(bg) if bg else None,
    }
    if with_bold:
        d['bold'] = bool(cell.font and cell.font.bold)
    return d

def normalize_ean_value(v):
    """Devuelve el EAN en su forma canonica de 13 digitos.

    openpyxl (data_only) lee algunas celdas EAN como numero en vez de texto, y
    al convertir a str un UPC-A pierde el 0 inicial ("0195..." -> "195...") o
    puede arrastrar un ".0" si vino como float. Si el valor mostrado difiere del
    nombre del fichero en disco (que sí lleva los 13 digitos con 0 delante) la
    foto no casa con el manifiesto y no carga. Normalizando aqui, en el origen,
    el valor que se pinta y el que va al catalogo coinciden siempre con el disco.
    Solo toca cadenas puramente numericas; cualquier otro valor se deja igual.
    """
    s = str(v).strip()
    if re.fullmatch(r'\d+\.0', s):   # float accidental: "195951406015.0"
        s = s[:-2]
    if s.isdigit() and len(s) == 12:  # UPC-A -> EAN-13 con 0 delante
        return '0' + s
    return s

def cell_style_attr(c):
    """style="..." inline (fondo+texto) para un dict de celda, o '' si no tiene color."""
    if not c['bg']:
        return ""
    fg = c['fg'] or '#111318'
    return f' style="background:{c["bg"]};color:{fg};"'

def slugify(s):
    s = s.lower()
    s = re.sub(r'[^a-z0-9]+', '-', s)
    return s.strip('-')

def sec_id_for(sheet_idx, si):
    return f'sec-{sheet_idx}-{si}'

def connectivity_label(cols):
    for c in cols:
        cl = c.strip().lower()
        if cl == 'lte':
            return 'LTE'
        if cl in ('wifi', 'wi-fi'):
            return 'Wi-Fi'
    return None

def clean_sheet_name(name):
    name = name.strip()
    name = name.replace('&amp;', '&').replace('&AMP;', '&')
    return name

def parse_lob_and_leaf(clean_name):
    if ' - ' in clean_name:
        lob, leaf = clean_name.split(' - ', 1)
        return lob.strip(), leaf.strip()
    return clean_name, clean_name

def panel_title_for(lob, leaf):
    if lob == leaf:
        return leaf
    if lob == 'EOL':
        return f'{leaf} EOL'
    if lob in ('iPad', 'iPhone', 'Apple Watch'):
        if leaf.lower().startswith(lob.lower()) or leaf.lower().startswith('pencil') or leaf.lower().startswith('strap'):
            return leaf
        return f'{lob} {leaf}'
    return leaf

def find_header_rows(ws, maxcol):
    header_rows = []
    # Filas que son una nota combinada a lo ancho (p.ej. "Requiere un iPhone
    # 11...", "Cellular solo funciona..."): nunca deben tratarse como
    # cabecera de una seccion nueva, aunque esten en negrita y la fila de
    # debajo tenga contenido (p.ej. otra nota justo a continuacion). Sin esto,
    # dos notas consecutivas hacen que la primera se confunda con una
    # cabecera y la segunda se trate como una fila de datos de una sección
    # fantasma, descuadrando los menus de navegacion.
    note_rows = set(
        mr.min_row for mr in ws.merged_cells.ranges
        if mr.min_row == mr.max_row and mr.min_col == 1 and mr.max_col > mr.min_col
    )
    for r in range(1, ws.max_row + 1):
        if r in note_rows:
            continue
        c1 = ws.cell(row=r, column=1)
        is_bold_1 = bool(c1.font and c1.font.bold)
        v1 = c1.value
        
        has_known = False
        for c in range(2, maxcol + 1):
            val = ws.cell(row=r, column=c).value
            if val is not None and str(val).strip().lower() in KNOWN_LABELS:
                has_known = True
                break
        
        # Case 1: Row 1 with empty c1 but known labels in other cols (e.g. iPhone DEMO)
        if r == 1 and is_bold_1 and (v1 is None or not str(v1).strip()) and has_known:
            header_rows.append(r)
            continue
            
        if not (is_bold_1 and v1 is not None and str(v1).strip()):
            continue
            
        # Case 2: c1 bold and has known labels in other cols
        if has_known:
            header_rows.append(r)
            continue
            
        # Case 3: c1 bold, other cols empty/headers, and next row is data
        other_vals = [ws.cell(row=r, column=c).value for c in range(2, maxcol + 1)]
        if all(v is None or str(v).strip().lower() in KNOWN_LABELS for v in other_vals):
            if r + 1 <= ws.max_row:
                next_c1 = ws.cell(row=r+1, column=1)
                next_c2 = ws.cell(row=r+1, column=2)
                if next_c1.value is not None or next_c2.value is not None:
                    header_rows.append(r)
    return header_rows

def _is_blank_row(ws, r, maxcol):
    return all(ws.cell(row=r, column=c).value is None for c in range(1, maxcol + 1))

def group_header_rows(ws, header_rows, maxcol):
    """Agrupa cabeceras consecutivas (p.ej. un modelo seguido de su propia
    cabecera de accesorios) en una sola tarjeta cuando entre ambas no hay
    ninguna fila completamente vacia. Una linea vacia en el Excel marca el
    final de todo lo que va junto con un producto, señalando que debe
    empezar la tarjeta siguiente; si no hay linea vacia, la cabecera
    siguiente se trata como una subseccion dentro de la misma tarjeta."""
    if not header_rows:
        return []
    groups = [[header_rows[0]]]
    for hr in header_rows[1:]:
        prev_hr = groups[-1][-1]
        blank_between = any(_is_blank_row(ws, r, maxcol) for r in range(prev_hr, hr))
        if blank_between:
            groups.append([hr])
        else:
            groups[-1].append(hr)
    return groups

def process_normal_sheet(ws, clean_name, lob, leaf):
    real_maxcol = 0
    for row in ws.iter_rows(max_row=min(ws.max_row, 400)):
        for c in row:
            if c.value is not None and c.column <= 10:
                real_maxcol = max(real_maxcol, c.column)
    maxcol = max(real_maxcol, 3)
    maxrow = ws.max_row

    header_rows = find_header_rows(ws, maxcol)

    synthetic_header = False
    if not header_rows:
        default_labels = ['Modelo', 'P/N', 'ECI', 'EAN']
        header_rows = [0]
        synthetic_header = True

    if synthetic_header:
        groups = [[0]]
    else:
        # Solo agrupamos cabeceras consecutivas en una misma tarjeta si la
        # hoja usa de verdad el convenio de linea vacia como separador. Si la
        # hoja no tiene ninguna fila vacia (la inmensa mayoria, todavia no
        # migradas a este convenio), cada cabecera sigue siendo su propia
        # tarjeta como hasta ahora, para no fusionar hojas enteras en una
        # unica tarjeta gigante por falta de ese marcador.
        sheet_has_blank = any(_is_blank_row(ws, r, maxcol) for r in range(1, maxrow + 1))
        if sheet_has_blank:
            groups = group_header_rows(ws, header_rows, maxcol)
        else:
            groups = [[hr] for hr in header_rows]

    # Mapa de combinaciones horizontales: (fila, col_inicio) -> col_fin.
    # Excel guarda el texto solo en la celda superior-izquierda del rango
    # combinado; el resto de celdas devuelven None. Sin esto, una nota a lo
    # ancho (p.ej. "Requiere un iPhone 11...") se pintaria solo en la primera
    # columna y las demas saldrian en blanco. Lo convertimos en un unico
    # <td colspan>.
    hmerge = {}
    for mr in ws.merged_cells.ranges:
        if mr.min_row == mr.max_row and mr.max_col > mr.min_col:
            hmerge[(mr.min_row, mr.min_col)] = mr.max_col

    default_labels = ['Modelo', 'P/N', 'ECI', 'EAN']

    sections = []
    for gi, group in enumerate(groups):
        subsecs = []
        for j, hr in enumerate(group):
            if synthetic_header:
                title = panel_title_for(lob, leaf)
                col_labels = default_labels[:maxcol] if maxcol <= len(default_labels) else default_labels + ['']*(maxcol-len(default_labels))
                data_start = 1
            else:
                raw_title = ws.cell(row=hr, column=1).value
                title = str(raw_title).strip() if raw_title is not None and str(raw_title).strip() else panel_title_for(lob, leaf)
                col_labels = []
                for c in range(1, maxcol+1):
                    v = ws.cell(row=hr, column=c).value
                    col_labels.append(str(v).strip() if v is not None else '')
                while col_labels and col_labels[-1] == '':
                    col_labels.pop()
                data_start = hr+1

            if j + 1 < len(group):
                sub_end = group[j+1] - 1
            elif gi + 1 < len(groups) and groups[gi+1][0] != 0:
                sub_end = groups[gi+1][0] - 1
            else:
                sub_end = maxrow

            ncols = len(col_labels)
            rows = []
            color_rows = []
            # Dentro de una misma subseccion, la fila "Accesorios recomendados"
            # (banda amarilla) puede aparecer como simple fila de datos en
            # negrita, sin volver a repetir cabeceras P/N/ECI/EAN -- por lo que
            # no siempre se detecta como cabecera propia mas arriba. Para no
            # depender de ese detalle, se corta aqui tambien a pie de fila: en
            # cuanto el texto de la columna 1 empieza por "accesorio"/"access",
            # esa fila y todas las siguientes de esta subseccion quedan fuera
            # del calculo de color de familia (aunque tengan EAN con relleno),
            # aunque siguen pintandose con normalidad en la tabla.
            past_accessory_marker = False
            for r in range(data_start, sub_end+1):
                rowvals = [ws.cell(row=r, column=c).value for c in range(1, ncols+1)]
                if all(v is None for v in rowvals):
                    continue
                c1v = ws.cell(row=r, column=1).value
                if c1v is not None:
                    c1low = str(c1v).strip().lower()
                    if c1low.startswith('access') or c1low.startswith('accesor'):
                        past_accessory_marker = True
                cells = []
                c = 1
                while c <= ncols:
                    cell = ws.cell(row=r, column=c)
                    d = cell_to_dict(cell)
                    # El EAN debe quedar siempre en su forma canonica de 13 digitos
                    # para casar con el nombre de la foto en disco y con el manifiesto.
                    if c-1 < len(col_labels) and col_labels[c-1].strip().upper() == 'EAN':
                        d['v'] = normalize_ean_value(d['v'])
                    span_end = hmerge.get((r, c))
                    if span_end:
                        span = min(span_end, ncols) - c + 1
                        if span > 1:
                            d['span'] = span
                            cells.append(d)
                            c += span
                            continue
                    cells.append(d)
                    c += 1
                rows.append(cells)
                if not past_accessory_marker:
                    color_rows.append(cells)

            if not rows:
                continue

            low_title = title.lower()
            sub_is_accessory = low_title.startswith('access') or low_title.startswith('accesor')
            subsecs.append({
                'title': title,
                'columns': col_labels,
                'rows': rows,
                'color_rows': [] if sub_is_accessory else color_rows,
                'is_accessory': sub_is_accessory,
            })

        if not subsecs:
            continue

        main = subsecs[0]
        all_rows = []
        for ss in subsecs:
            all_rows.extend(ss['rows'])

        sections.append({
            'title': main['title'],
            'columns': main['columns'],
            'rows': all_rows,
            'color_rows': main['color_rows'],
            'is_accessory': main['is_accessory'],
            'count': len(all_rows),
            'groups': subsecs if len(subsecs) > 1 else None,
        })

    return sections

def count_unique_references(sheets_data):
    """Cuenta referencias para el total del landing, sin duplicar un mismo
    producto (mismo EAN) que aparece listado en varias hojas a la vez — por
    ejemplo un cable o cargador comun que se repite como accesorio en cada
    modelo de iPhone/iPad. Las filas sin EAN (notas combinadas, hojas tipo
    matriz de compatibilidad) se cuentan igual que antes, 1 a 1."""
    seen_eans = set()
    total = 0
    for s in sheets_data:
        if s['type'] == 'matrix':
            total += s['count']
            continue
        for sec in s['sections']:
            cols = sec['columns']
            ean_idx = next((i for i, c in enumerate(cols) if c.strip().upper() == 'EAN'), None)
            for row in sec['rows']:
                ean_val = None
                if ean_idx is not None:
                    col_pos = 0
                    for c in row:
                        span = c.get('span') or 1
                        if col_pos <= ean_idx < col_pos + span:
                            ean_val = c['v']
                            break
                        col_pos += span
                if ean_val:
                    ev = ean_val.strip()
                    if ev in seen_eans:
                        continue
                    seen_eans.add(ev)
                total += 1
    return total

def process_matrix_sheet(ws):
    maxcol = ws.max_column
    maxrow = ws.max_row
    data = []
    col_count = 0
    for row in ws.iter_rows(min_row=1, max_row=maxrow, max_col=maxcol):
        vals = [c.value for c in row]
        while vals and vals[-1] is None:
            vals.pop()
        col_count = max(col_count, len(vals))
    for r in range(1, maxrow+1):
        rowvals = []
        any_val = False
        for c in range(1, col_count+1):
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if v is not None:
                any_val = True
            rowvals.append(cell_to_dict(cell, with_bold=True))
        if any_val:
            data.append(rowvals)
    return {'rows': data, 'col_count': col_count}

def render_sections(sheet):
    parts = []
    title_counts = {}
    for sec in sheet['sections']:
        if sec['is_accessory']:
            continue
        title_counts[sec['title']] = title_counts.get(sec['title'], 0) + 1

    family_grads = sheet.get('family_grads') or {}

    # Empareja cada tarjeta de accesorios con el modelo de producto justo
    # anterior en la hoja, para que herede su mismo color de familia (cabecera,
    # borde y triangulo) y para poder sincronizar su apertura/cierre en el JS.
    pair_map = {}
    pending_product_si = None
    for si, sec in enumerate(sheet['sections']):
        if sec['is_accessory']:
            if pending_product_si is not None:
                pair_map[si] = pending_product_si
                pair_map[pending_product_si] = si
                pending_product_si = None
        else:
            pending_product_si = si

    for si, sec in enumerate(sheet['sections']):
        cols = sec['columns']
        mono_idx = set(i for i, c in enumerate(cols) if c.strip().lower() in MONO_LABELS)
        acc_class = " accessory-card" if sec['is_accessory'] else ""
        sec_id = sec_id_for(sheet['idx'], si)
        display_title = sec['title']
        if title_counts.get(display_title, 0) > 1:
            conn = connectivity_label(cols)
            if conn:
                display_title = f"{display_title} ({conn})"
        partner_si = pair_map.get(si)
        if sec['is_accessory'] and partner_si is not None:
            grad, wash = family_grads.get(partner_si, (None, None))
        else:
            grad, wash = family_grads.get(si, (None, None))
        fam_class = " fam-card" if grad else ""
        fam_style = fam_style_attr(grad, wash)
        pair_attr = ''
        if partner_si is not None:
            partner_id = sec_id_for(sheet['idx'], partner_si)
            pair_attr = f' data-pair-card="{partner_id}-card"'
        parts.append(f'<div class="card{acc_class}{fam_class}" id="{sec_id}-card" data-section{fam_style}{pair_attr}>')
        parts.append(f'  <div class="card-header" data-toggle="{sec_id}">')
        parts.append(f'    <div class="card-title-wrap">')
        parts.append(f'      <span class="chevron">▶</span>')
        title_cls = ' card-title-accessory' if sec['is_accessory'] else ''
        parts.append(f'      <h3 class="card-title{title_cls}">{esc(display_title)}</h3>')
        parts.append(f'    </div>')
        parts.append(f'  </div>')
        parts.append(f'  <div class="card-body" id="{sec_id}">')
        groups = sec.get('groups') or [sec]
        for gi, grp in enumerate(groups):
            grp_cols = grp['columns']
            grp_mono_idx = set(i for i, c in enumerate(grp_cols) if c.strip().lower() in MONO_LABELS)
            if gi > 0:
                parts.append(f'    <div class="group-divider"><span class="group-divider-text">{esc(grp["title"])}</span></div>')
            parts.append(f'    <div class="table-wrap"><table>')
            parts.append('      <thead><tr>')
            for c in grp_cols:
                parts.append(f'<th>{esc(c)}</th>')
            parts.append('      </tr></thead>')
            parts.append('      <tbody>')
            for row in grp['rows']:
                row_text = " ".join(c['v'] for c in row).lower()
                c1v = row[0]['v'] if row else ''
                is_acc_marker = bool(c1v) and c1v.strip().lower().startswith(('accesor', 'access'))
                tr_cls = ' class="accessory-marker-row"' if is_acc_marker else ''
                parts.append(f'<tr{tr_cls} data-search="{esc(row_text)}">')
                col_pos = 0
                for c in row:
                    style = cell_style_attr(c)
                    cls = ' class="mono"' if col_pos in grp_mono_idx else ''
                    val = esc(c['v']) if c['v'] else ""
                    span = c.get('span')
                    span_attr = f' colspan="{span}"' if span else ''
                    parts.append(f'<td{cls}{style}{span_attr}><span class="cell-text">{val}</span></td>')
                    col_pos += span if span else 1
                parts.append('</tr>')
            parts.append('      </tbody>')
            parts.append('    </table></div>')
        parts.append('  </div>')
        parts.append('</div>')
    return "\n".join(parts)

def render_matrix(sheet):
    rows = sheet['matrix']['rows']
    parts = []
    parts.append('<div class="card matrix-card">')
    parts.append('  <div class="table-wrap"><table class="matrix-table">')
    parts.append('    <tbody>')
    for row in rows:
        any_val = any(c['v'] for c in row)
        if not any_val:
            parts.append('<tr class="matrix-spacer"><td colspan="99"></td></tr>')
            continue
        parts.append('<tr>')
        for c in row:
            style = cell_style_attr(c)
            cls = 'th-cell' if c.get('bold') else 'td-cell'
            val = esc(c['v'])
            parts.append(f'<td class="{cls}"{style}><span class="cell-text">{val}</span></td>')
        parts.append('</tr>')
    parts.append('    </tbody>')
    parts.append('  </table></div>')
    parts.append('</div>')
    return "\n".join(parts)

PICTURES_DIR = os.path.join(BASE_DIR, 'pictures')
IMAGES_JSON_PATH = os.path.join(BASE_DIR, 'images.json')

# Nombres que macOS / iCloud / los navegadores generan al bajar dos veces el
# mismo fichero: "<ean> 2.webp", "<ean> 3.webp", "<ean> copia.webp",
# "<ean> copia 2.webp", "<ean> copy.webp". Definida en img_naming.py y
# compartida con resize_inplace.py: los dos scripts deben estar de acuerdo
# sobre que es un duplicado o arriesgan borrar/conservar fotos de forma
# inconsistente entre el redimensionado y el build.


def canonical_ean_stem(stem, catalog_eans):
    """Devuelve el nombre canonico para un stem de imagen.

    Los EAN de 12 digitos (UPC-A) se guardan en la base de datos con un 0
    delante para ser EAN-13. Si en disco aparece la version corta pero el
    catalogo solo conoce la larga, se corrige. La condicion es deliberadamente
    estricta: solo se renombra cuando el nombre con 0 EXISTE de verdad en el
    catalogo, de modo que nunca se inventa una referencia.
    """
    if not stem.isdigit() or not catalog_eans:
        return stem
    if stem in catalog_eans:
        return stem
    if len(stem) == 12 and ('0' + stem) in catalog_eans:
        return '0' + stem
    return stem


def clean_pictures_dir(catalog_eans=None):
    """Limpia pictures/ antes de construir el manifiesto.

    Tres pasadas, en este orden:

    1. Duplicados ("<ean> 2.webp"). Si la imagen original existe, el duplicado
       se BORRA. Si no existe (el duplicado es la unica foto del producto), se
       RESCATA renombrandolo al nombre bueno en vez de perder la imagen.
    2. Normalizacion de EAN: "<12 digitos>.webp" -> "<0+12 digitos>.webp"
       cuando el catalogo usa la forma larga. Asi una foto nueva descargada con
       el nombre antiguo se coloca sola, sin renombrados manuales.
    3. Carpetas vacias: se dejan intactas (puede que el usuario las acabe de
       crear para rellenarlas); solo se informa.

    Devuelve un dict con el detalle de lo hecho para poder imprimir el resumen.
    """
    report = {'deleted': [], 'rescued': [], 'normalized': [], 'skipped': [], 'empty_folders': []}
    if not os.path.isdir(PICTURES_DIR):
        return report

    for folder in sorted(os.listdir(PICTURES_DIR)):
        folder_path = os.path.join(PICTURES_DIR, folder)
        if not os.path.isdir(folder_path):
            continue

        names = [f for f in os.listdir(folder_path) if f.lower().endswith('.webp')]
        stems = set(n[:-5] for n in names)

        # --- 1) Duplicados ---
        for fname in sorted(names):
            stem = fname[:-5]
            m = DUP_STEM_RE.match(stem)
            if not m:
                continue
            base = m.group('base').strip()
            if not base.isdigit():
                # Nombre raro que no sabemos reconstruir: no se toca.
                report['skipped'].append((folder, fname))
                continue
            canon = canonical_ean_stem(base, catalog_eans)
            src = os.path.join(folder_path, fname)
            if canon in stems or base in stems:
                os.remove(src)
                stems.discard(stem)
                report['deleted'].append((folder, fname))
            else:
                dst_name = canon + '.webp'
                os.rename(src, os.path.join(folder_path, dst_name))
                stems.discard(stem)
                stems.add(canon)
                report['rescued'].append((folder, fname, dst_name))

        # --- 2) Normalizacion de EAN de 12 digitos ---
        for stem in sorted(s for s in stems if s.isdigit()):
            canon = canonical_ean_stem(stem, catalog_eans)
            if canon == stem:
                continue
            src = os.path.join(folder_path, stem + '.webp')
            dst = os.path.join(folder_path, canon + '.webp')
            if canon in stems:
                # Ya existe la version correcta: el nombre corto sobra.
                os.remove(src)
                report['deleted'].append((folder, stem + '.webp'))
            else:
                os.rename(src, dst)
                stems.add(canon)
                report['normalized'].append((folder, stem + '.webp', canon + '.webp'))

        if not stems:
            report['empty_folders'].append(folder)

    return report


def diff_image_manifest(new_manifest):
    """Compara el manifiesto recien escaneado con el images.json anterior.

    Sirve para dejar constancia en el log de que carpetas e imagenes son
    NUEVAS respecto al build previo (y cuales han desaparecido). El manifiesto
    se regenera siempre desde disco, asi que esto no cambia el resultado: solo
    hace visible lo que se ha detectado.
    """
    try:
        with open(IMAGES_JSON_PATH, encoding='utf-8') as f:
            old = json.load(f)
    except Exception:
        old = {}
    old_sets = {k: set(v) for k, v in old.items()}
    new_sets = {k: set(v) for k, v in new_manifest.items()}

    new_folders = sorted(set(new_sets) - set(old_sets))
    gone_folders = sorted(set(old_sets) - set(new_sets))
    added, removed = [], []
    for folder in sorted(new_sets):
        for ean in sorted(new_sets[folder] - old_sets.get(folder, set())):
            added.append((folder, ean))
    for folder in sorted(old_sets):
        for ean in sorted(old_sets[folder] - new_sets.get(folder, set())):
            removed.append((folder, ean))
    return {
        'new_folders': new_folders,
        'gone_folders': gone_folders,
        'added': added,
        'removed': removed,
        'first_run': not old,
    }


def build_image_manifest():
    """Escanea pictures/<hoja>/<ean>.webp y devuelve un dict {hoja: [ean, ...]}
    con TODAS las imagenes .webp presentes en disco en el momento del build.

    Se ejecuta en cada build, asi que si se anaden o quitan fotos (o cambian
    referencias en la base de datos), el manifiesto se rehace automaticamente
    sin intervencion manual. El front lo usa para saber, sin tocar la red, que
    referencias tienen foto: las que no aparecen aqui no disparan ninguna
    peticion. Solo .webp — no existen ni se contemplan .jpeg.

    El nombre de una imagen valida es ESTRICTAMENTE un EAN (solo digitos) mas
    ".webp". Cualquier otro fichero se ignora: en particular los duplicados que
    genera el bucle contra la nube con el sufijo " 2" ("<ean> 2.webp"), pero
    tambien cualquier nombre no numerico que pudiera aparecer en el futuro.
    Aceptar solo lo numerico (en vez de rechazar patrones concretos) mantiene el
    manifiesto limpio por defecto ante nombres inesperados.
    """
    manifest = {}
    if not os.path.isdir(PICTURES_DIR):
        return manifest
    for folder in os.listdir(PICTURES_DIR):
        folder_path = os.path.join(PICTURES_DIR, folder)
        if not os.path.isdir(folder_path):
            continue
        eans = []
        for fname in os.listdir(folder_path):
            if not fname.lower().endswith('.webp'):
                continue
            stem = fname[:-5]  # quita ".webp"
            # Solo nombres que son estrictamente un EAN (digitos y nada mas).
            if not stem.isdigit():
                continue
            eans.append(stem)
        if eans:
            manifest[folder] = sorted(eans)
    return manifest

def main():
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    mod = wb.properties.modified or wb.properties.created
    last_updated = mod.strftime('%d/%m/%Y') if mod else datetime.date.today().strftime('%d/%m/%Y')

    sheets_data = []
    for i, name in enumerate(wb.sheetnames):
        ws = wb[name]
        clean = clean_sheet_name(name)
        lob, leaf = parse_lob_and_leaf(clean)
        display_title = panel_title_for(lob, leaf)
        if 'pencil' in clean.lower() and 'keyboard' in clean.lower():
            matrix = process_matrix_sheet(ws)
            sheets_data.append({
                'idx': i,
                'name': clean,
                'lob': lob,
                'leaf': leaf,
                'title': display_title,
                'slug': slugify(clean) + f'-{i}',
                'type': 'matrix',
                'matrix': matrix,
                'count': len(matrix['rows']),
            })
        else:
            sections = process_normal_sheet(ws, clean, lob, leaf)
            total = sum(s['count'] for s in sections)
            sheets_data.append({
                'idx': i,
                'name': clean,
                'lob': lob,
                'leaf': leaf,
                'title': display_title,
                'slug': slugify(clean) + f'-{i}',
                'type': 'sections',
                'sections': sections,
                'count': total,
            })

    # Group sheets by LOB while preserving the order they appear in the Excel workbook
    lobs = {}
    lob_order = []
    for s in sheets_data:
        lob_name = s['lob']
        if lob_name not in lobs:
            lobs[lob_name] = []
            lob_order.append(lob_name)
        lobs[lob_name].append(s)

    # Flat searchable rows
    all_rows = []
    for s in sheets_data:
        if s['type'] != 'sections':
            continue
        for sec in s['sections']:
            cols = sec['columns']
            mono_idx = set(i for i, c in enumerate(cols) if c.strip().lower() in MONO_LABELS)
            for row in sec['rows']:
                cells = []
                search_parts = [s['name'], s['title'], sec['title']]
                col_pos = 0
                for c in row:
                    label = cols[col_pos] if col_pos < len(cols) else ''
                    cell = {
                        'label': label,
                        'v': c['v'],
                        'bg': c['bg'],
                        'fg': c['fg'],
                        'mono': col_pos in mono_idx,
                    }
                    if c.get('span'):
                        cell['span'] = c['span']
                    cells.append(cell)
                    if c['v']:
                        search_parts.append(c['v'])
                    col_pos += c.get('span') or 1
                all_rows.append({
                    'sheet': s['title'],
                    'sheetFolder': s['name'],
                    'slug': s['slug'],
                    'section': sec['title'],
                    'isAccessory': sec['is_accessory'],
                    'cells': cells,
                    'search': ' '.join(search_parts).lower(),
                    # Forma solo-digitos de cada referencia numerica de la fila
                    # (ECI, EAN...), con un espacio entre referencias para que una
                    # busqueda no cruce de una a otra. Permite encontrar una
                    # referencia tecleada con OTRA agrupacion de puntos -- o sin
                    # ninguno -- ya que las etiquetas de ECI no siempre parten los
                    # 11 digitos igual (a veces hay un punto de mas). La cadena
                    # 'search' de arriba conserva el formato original con puntos;
                    # esta es el respaldo normalizado.
                    'searchnum': ' '.join(
                        re.sub(r'\D', '', p) for p in search_parts
                        if len(re.sub(r'\D', '', p)) >= 6
                    ),
                })

    # Build Level 1 and Level 2 navigation
    # El reseteo a Inicio ya no vive en un boton casa: lo dispara la zona de
    # marca (arcos + ECIarc), que es clicable. Asi se ahorra el espacio del boton.
    level1_btns = []
    level2_rows = []

    for lob in lob_order:
        sheets_in_lob = lobs[lob]
        if len(sheets_in_lob) == 1:
            s = sheets_in_lob[0]
            level1_btns.append(
                f'<button type="button" class="nav-btn" data-target="{s["slug"]}">'
                f'{esc(s["title"])}</button>'
            )
            continue
        lob_slug = slugify(lob)
        total_count = sum(s['count'] for s in sheets_in_lob)
        level1_btns.append(
            f'<button type="button" class="nav-btn" data-lob="{lob_slug}">'
            f'{esc(lob)}</button>'
        )
        leaf_btns = []
        # Mismo coloreo por "familia" que el submenu de modelos, aplicado ahora
        # al menu de hojas del LOB (nivel 2). La paleta de cada hoja es la union
        # de los colores de producto de sus secciones; luego se agrupan las
        # hojas por familia (18 Pro / 18 Pro Max comparten paleta, etc.) para
        # que hermanas de distinto tamano salgan con el mismo degradado.
        def _sheet_colors(sh):
            if sh.get('type') != 'sections':
                return []
            return _union_colors(
                [section_product_colors(sec) for sec in sh['sections'] if not sec['is_accessory']]
            )
        leaf_members = {}
        leaf_order = []
        for s in sheets_in_lob:
            k = family_key(s['leaf'])
            if k not in leaf_members:
                leaf_members[k] = []
                leaf_order.append(k)
            leaf_members[k].append(s)
        leaf_fam_palette = {
            k: _union_colors([_sheet_colors(s) for s in leaf_members[k]]) for k in leaf_order
        }
        for s in sheets_in_lob:
            pal = leaf_fam_palette[family_key(s['leaf'])]
            # Se guarda en la propia hoja para que, al pintar su panel mas
            # abajo, las tarjetas de una hoja de familia unica hereden la MISMA
            # paleta del LOB (la del boton de nivel 2) y el panel entero sea
            # coherente con el menu.
            s['leaf_palette'] = pal
            grad, wash = gradient_vars(pal)
            fam_cls = ' nav-btn-fam' if grad else ''
            leaf_btns.append(
                f'<button type="button" class="nav-btn nav-btn-leaf{fam_cls}"'
                f'{fam_style_attr(grad, wash)} data-target="{s["slug"]}">'
                f'{esc(s["leaf"])}</button>'
            )
        level2_rows.append(
            f'<div class="btn-row level2" id="lobrow-{lob_slug}">{"".join(leaf_btns)}</div>'
        )

    total_sheets = len(sheets_data)
    total_refs_all = count_unique_references(sheets_data)
    update_html = f'<span class="home-updated">Última actualización: <strong>{esc(last_updated)}</strong></span>' if last_updated else ''

    home_panel = f'''
      <section class="panel active" id="home" data-panel-name="Inicio">
        <div class="panel-header home-header">
          <div class="home-today">
            <span class="home-today-date" id="homeTodayDate"></span>
          </div>
          <div class="home-links">
            <a href="#" id="homeHelpLink" class="home-link home-link-primary"><span class="home-link-icon">❓</span> Ayuda — cómo usar la página</a>
            <a href="https://pjexposito.github.io/planmy.html" target="_blank" rel="noopener noreferrer" class="home-link"><span class="home-link-icon">🧮</span> Calculadora PlanMy</a>
            <a href="#" id="homeIncidentLink" class="home-link"><span class="home-link-icon">✉️</span> Reportar una incidencia</a>
          </div>
          <div class="home-stats">
            <span><strong>{total_refs_all}</strong> referencias</span>
            {update_html}
          </div>
        </div>
      </section>'''

    search_results_panel = '''
      <section class="panel" id="search-results" data-panel-name="Resultados">
        <div class="panel-header">
          <h2>Resultados de búsqueda</h2>
          <p class="search-results-count" id="searchResultsCount"></p>
        </div>
        <div class="table-wrap"><table id="searchResultsTable">
          <tbody id="searchResultsBody"></tbody>
        </table></div>
      </section>'''

    panels = [home_panel, search_results_panel]
    for s in sheets_data:
        if s['type'] == 'sections':
            model_secs = [(si, sec) for si, sec in enumerate(s['sections']) if not sec['is_accessory']]
            # Colorear por "familia" de producto. Dos criterios, en este orden:
            #  - Si la hoja es UNA sola familia (p.ej. iPad A16 = Wifi/LTE, o
            #    iPhone 18 Pro Max) y cuelga de un LOB multi-hoja, todas sus
            #    tarjetas toman el color de familia del LOB (el del boton de
            #    nivel 2). Asi el panel entero es de un color y las variaciones
            #    hermanas (18 Pro / 18 Pro Max) coinciden entre si.
            #  - Si la hoja tiene VARIAS familias internas (p.ej. Apple Watch
            #    Series 12, por material; o la hoja Mac con Neo/Air/Pro/iMac...)
            #    se colorea cada una con su tono para distinguirlas.
            # Se calcula ANTES de render_sections para que el mismo tono pinte
            # a la vez el boton del submenu y la tarjeta a la que apunta.
            s['family_grads'] = {}
            if model_secs:
                sec_list = [sec for _, sec in model_secs]
                within_pal, n_families = family_palettes_for(sec_list)
                leaf_palette = s.get('leaf_palette')
                if leaf_palette and n_families == 1:
                    # Hoja de familia unica dentro de un LOB multi-hoja: todas las
                    # tarjetas heredan la paleta del boton de nivel 2 (18 Pro /
                    # 18 Pro Max y demas hermanas salen identicas entre si).
                    grad, wash = gradient_vars(leaf_palette)
                    if grad:
                        s['family_grads'] = {si: (grad, wash) for si, _ in model_secs}
                elif len(model_secs) > 1:
                    # Hoja con varias familias internas: cada una con la paleta de
                    # sus propios colores de producto para distinguirlas.
                    for pos, (si, _) in enumerate(model_secs):
                        grad, wash = gradient_vars(within_pal[pos])
                        if grad:
                            s['family_grads'][si] = (grad, wash)

            body = render_sections(s)
            toolbar = ''
            model_row = ''
            if len(model_secs) > 1:
                title_counts = {}
                for _, sec in model_secs:
                    title_counts[sec['title']] = title_counts.get(sec['title'], 0) + 1
                model_btns = []
                for si, sec in model_secs:
                    sec_id = sec_id_for(s['idx'], si)
                    label = sec['title']
                    if title_counts[label] > 1:
                        conn = connectivity_label(sec['columns'])
                        if conn:
                            label = f"{label} ({conn})"
                    grad, wash = s['family_grads'].get(si, (None, None))
                    fam_cls = ' nav-btn-fam' if grad else ''
                    model_btns.append(
                        f'<button type="button" class="nav-btn nav-btn-model{fam_cls}"'
                        f'{fam_style_attr(grad, wash)} data-model-target="{sec_id}-card">'
                        f'{esc(label)}</button>'
                    )
                model_row = f'<div class="btn-row level3">{"".join(model_btns)}</div>'
        else:
            body = render_matrix(s)
            toolbar = '<div class="panel-toolbar"><p class="matrix-hint">Tabla de compatibilidad Apple Pencil / Magic Keyboard × modelo de iPad.</p></div>'
            model_row = ''

        panels.append(f'''
      <section class="panel" id="{s['slug']}" data-panel-name="{esc(s['title'])}" data-sheet-folder="{esc(s['name'])}">
        <div class="panel-header">
          <h2>{esc(s['title'])}</h2>
        </div>
        {model_row}
        {toolbar}
        <div class="panel-content" data-panel-content="{s['slug']}">
          {body}
        </div>
      </section>''')

    all_rows_json = json.dumps(all_rows, ensure_ascii=False)

    # --- Write external files ---
    css_path = os.path.join(BASE_DIR, 'style.css')
    with open(css_path, 'w', encoding='utf-8') as f:
        f.write(CSS_CONTENT)
    print(f"CSS generado: {css_path}")

    data_path = os.path.join(BASE_DIR, 'data.json')
    with open(data_path, 'w', encoding='utf-8') as f:
        f.write(all_rows_json)
    print(f"Datos generados: {data_path} ({len(all_rows)} filas)")

    # --- Image manifest: qué referencias tienen foto .webp en disco ---
    # 1) Limpieza previa de pictures/: duplicados "<ean> 2.webp" y nombres de
    #    EAN sin el 0 inicial. Se le pasa el conjunto de EAN reales del catálogo
    #    para que ningún renombrado invente una referencia inexistente.
    catalog_eans = set()
    for row in all_rows:
        for c in row['cells']:
            if (c.get('label') or '').strip().upper() == 'EAN':
                v = (c.get('v') or '').strip()
                if v.isdigit():
                    catalog_eans.add(v)

    clean = clean_pictures_dir(catalog_eans)
    if clean['deleted'] or clean['rescued'] or clean['normalized'] or clean['skipped']:
        print("Limpieza de pictures/:")
        if clean['deleted']:
            print(f"   duplicados borrados: {len(clean['deleted'])}")
        if clean['rescued']:
            print(f"   duplicados rescatados (eran la unica foto): {len(clean['rescued'])}")
            for folder, old, new in clean['rescued'][:10]:
                print(f"      {folder}/{old} -> {new}")
        if clean['normalized']:
            print(f"   EAN normalizados a 13 digitos: {len(clean['normalized'])}")
            for folder, old, new in clean['normalized'][:10]:
                print(f"      {folder}/{old} -> {new}")
        if clean['skipped']:
            print(f"   ignorados (nombre no reconocible): {len(clean['skipped'])}")
            for folder, fname in clean['skipped'][:10]:
                print(f"      {folder}/{fname}")
    else:
        print("Limpieza de pictures/: nada que limpiar")

    # 2) Escaneo del arbol ya limpio y diff contra el build anterior, para que
    #    quede claro que carpetas/imagenes nuevas se han detectado.
    image_manifest = build_image_manifest()
    diff = diff_image_manifest(image_manifest)
    if diff['first_run']:
        print("Deteccion de novedades: primer manifiesto (no hay anterior con que comparar)")
    elif diff['new_folders'] or diff['added'] or diff['gone_folders'] or diff['removed']:
        print("Novedades detectadas en pictures/:")
        if diff['new_folders']:
            print(f"   carpetas nuevas: {len(diff['new_folders'])} -> {', '.join(diff['new_folders'])}")
        if diff['added']:
            print(f"   imagenes nuevas: {len(diff['added'])}")
            for folder, ean in diff['added'][:15]:
                print(f"      + {folder}/{ean}.webp")
            if len(diff['added']) > 15:
                print(f"      ... y {len(diff['added']) - 15} mas")
        if diff['gone_folders']:
            print(f"   carpetas desaparecidas: {', '.join(diff['gone_folders'])}")
        if diff['removed']:
            print(f"   imagenes que ya no estan: {len(diff['removed'])}")
            for folder, ean in diff['removed'][:15]:
                print(f"      - {folder}/{ean}.webp")
            if len(diff['removed']) > 15:
                print(f"      ... y {len(diff['removed']) - 15} mas")
    else:
        print("Novedades detectadas en pictures/: ninguna (sin cambios desde el build anterior)")

    orphans = sum(1 for folder, eans in image_manifest.items()
                  for e in eans if e not in catalog_eans)
    if orphans:
        print(f"   AVISO: {orphans} imagenes sin EAN correspondiente en Database.xlsx")

    images_path = os.path.join(BASE_DIR, 'images.json')
    with open(images_path, 'w', encoding='utf-8') as f:
        f.write(json.dumps(image_manifest, ensure_ascii=False))
    total_images = sum(len(v) for v in image_manifest.values())
    print(f"Manifiesto de imagenes generado: {images_path} "
          f"({total_images} imagenes en {len(image_manifest)} hojas)")

    # --- Write HTML (now lightweight) ---
    term_syn, term_max_words, term_report = load_term_synonyms()
    print(term_report)

    html_doc = TEMPLATE.replace('__LEVEL1_BTNS__', "".join(level1_btns)) \
                        .replace('__LEVEL2_ROWS__', "\n".join(level2_rows)) \
                        .replace('__TERM_SYNONYMS__', json.dumps(term_syn, ensure_ascii=False, sort_keys=True)) \
                        .replace('__TERM_MAX_WORDS__', str(term_max_words)) \
                        .replace('__PANELS__', "\n".join(panels))

    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html_doc)
    print(f"HTML generado exitosamente en: {HTML_PATH}")
    print(f"Total hojas: {total_sheets}, Total referencias: {total_refs_all}")

CSS_CONTENT = r"""/* ECIarc — El Corte Inglés Apple Reference Catalog — style.css */
:root{
  --bg: #0d0f14;
  --bg-alt: #131620;
  --panel: #171a24;
  --panel-2: #1d2130;
  --border: #2a2f3f;
  --text: #e8eaf0;
  --text-dim: #9aa1b5;
  --accent: #3ea25a;
  --accent-2: #2f8a4a;
  --accessory-bg: #1a2a1f;
  --accessory-border: #3ea25a;
  --radius: 12px;
  --mono: 'SF Mono', 'Menlo', 'Consolas', monospace;
}
body.light-mode{
  --bg: #f4f7f5;
  --bg-alt: #ffffff;
  --panel: #ffffff;
  --panel-2: #eef4f0;
  --border: #d9e3dc;
  --text: #14161c;
  --text-dim: #57614f;
  --accent: #0b7a3e;
  --accent-2: #0f9a52;
  --accessory-bg: #eaf7ee;
  --accessory-border: #0f9a52;
}
body.light-mode mark{background:#ffe067;color:#1a1a1a;}
*{box-sizing:border-box;}
html,body{margin:0;padding:0;}
html{ -webkit-text-size-adjust:100%; }
body{
  background: var(--bg);
  color: var(--text);
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  min-height:100vh;
  min-height:100dvh;
  -webkit-tap-highlight-color: transparent;
  overscroll-behavior-y: contain;
}
a{color:inherit;text-decoration:none;}
.topbar{position:sticky;top:0;z-index:50;background: var(--bg-alt);border-bottom:1px solid var(--border);padding-top: env(safe-area-inset-top, 0);padding-left: env(safe-area-inset-left, 0);padding-right: env(safe-area-inset-right, 0);}
.topbar-header{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:10px 14px;}
.brand-lockup{display:flex;align-items:flex-end;gap:3px;min-width:0;cursor:pointer;-webkit-tap-highlight-color:transparent;user-select:none;}
.brand-lockup:focus-visible{outline:2px solid var(--accent);outline-offset:3px;border-radius:6px;}
.brand-arc-mark{display:block;flex-shrink:0;width:40px;height:23px;}
.arc-stop-start{stop-color:#8E8E93;}
.arc-stop-end{stop-color:#2CA55E;}
body.light-mode .arc-stop-start{stop-color:#3A3A3C;}
body.light-mode .arc-stop-end{stop-color:#1E7145;}
.topbar-header h1{font-size:20px;margin:0;font-weight:700;line-height:.85;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.topbar-header h1 .brand-eci{color:#0b9444;}
body.light-mode .topbar-header h1 .brand-eci{color:#0b7a3e;}
.topbar-header h1 .brand-colon{background:linear-gradient(90deg,#0b9444,var(--text-dim));-webkit-background-clip:text;background-clip:text;color:transparent;}
body.light-mode .topbar-header h1 .brand-colon{background:linear-gradient(90deg,#0b7a3e,var(--text-dim));-webkit-background-clip:text;background-clip:text;}
.topbar-header h1 .brand-arc{color: var(--text-dim);}
.topbar-actions{display:flex;align-items:center;gap:6px;flex-shrink:0;}
.icon-btn{width:38px;height:38px;border-radius:9px;border:1px solid var(--border);background: var(--panel);color: var(--text);font-size:15px;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;transition: border-color .12s, background .12s;-webkit-tap-highlight-color: transparent;touch-action: manipulation;}
.icon-btn:hover{border-color: var(--accent);background: var(--panel-2);}
.icon-btn:active{transform: scale(.94);}
.icon-btn:disabled{opacity:.4;cursor:default;}
.icon-btn:disabled:active{transform:none;}
.icon-btn-font{font-size:12px;font-weight:700;gap:1px;}
.font-btn-sign{font-size:9px;position:relative;top:-3px;}
.search-bar-row{max-height:0;overflow:hidden;transition: max-height .18s ease, padding .18s ease;padding:0 14px;}
.search-bar-row.open{max-height:60px;padding:0 14px 10px;}
.search-input-wrap{position:relative;display:flex;align-items:center;}
.search-bar-row input{width:100%;padding:11px 36px 11px 12px;border-radius:10px;border:1px solid var(--border);background: var(--panel);color: var(--text);font-size:16px;outline:none;-webkit-appearance:none;appearance:none;}
.search-bar-row input:focus{border-color: var(--accent);}
.search-clear-btn{position:absolute;right:6px;top:50%;transform:translateY(-50%);width:26px;height:26px;border:none;border-radius:50%;background:transparent;color:var(--muted, #888);font-size:20px;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;padding:0;}
.search-clear-btn:hover{background: var(--border);color: var(--text);}
.btn-row{display:flex;flex-wrap:wrap;align-items:center;gap:6px;padding:9px 14px;row-gap:8px;}
.btn-row.level1{border-top:1px solid var(--border);}
.btn-row.level2{display:none;padding:0 14px;flex-wrap:wrap;}
.btn-row.level2.open{display:flex;padding:9px 14px;border-top:1px solid var(--border);background: rgba(255,255,255,.015);}
.btn-row.level3{padding:0 0 4px;margin:0 0 14px;border-bottom:1px solid var(--border);}
.nav-btn{flex-shrink:0;display:inline-flex;align-items:center;gap:6px;padding:9px 13px;min-height:38px;border-radius:20px;border:1px solid var(--border);background: var(--panel);color: var(--text-dim);font-size:13px;font-weight:600;cursor:pointer;transition: background .12s, color .12s, border-color .12s;white-space:nowrap;-webkit-tap-highlight-color: transparent;touch-action: manipulation;}
.nav-btn:hover{border-color: var(--accent);color: var(--text);}
.nav-btn:active{transform: scale(.96);}
.nav-btn.active{background: linear-gradient(90deg, var(--accent), var(--accent-2));color:#fff;border-color: transparent;}
.nav-btn-leaf{font-size:12.5px;padding:8px 12px;min-height:36px;}
.nav-btn-model{font-size:12.5px;padding:8px 12px;min-height:34px;background: var(--panel-2);}
/* Colorear por "familia" de producto. La paleta ya NO es un tono HSL fijo:
   son los colores reales de las celdas de las tarjetas (--fam-grad, lista de
   paradas; --fam-wash, esas mismas paradas a baja opacidad), calculados en el
   build desde Database.xlsx. Un MacBook Pro sale con 2 colores, un Neo con 4.
   Marco con degradado via padding-box/border-box (compatible Safari/iOS): la
   capa padding-box pinta el fondo del boton, la border-box el degradado del
   borde. El mismo degradado se repite en el boton del submenu y en el borde de
   la tarjeta a la que apunta, para percibirlos como el mismo grupo. */
.nav-btn-fam{border:1.5px solid transparent;background:
    linear-gradient(var(--panel-2),var(--panel-2)) padding-box,
    linear-gradient(135deg, var(--fam-grad)) border-box;}
.nav-btn-fam:hover{color: var(--text);background:
    linear-gradient(135deg, var(--fam-wash)) padding-box,
    linear-gradient(var(--panel-2),var(--panel-2)) padding-box,
    linear-gradient(135deg, var(--fam-grad)) border-box;}
/* Al activarse, el boton de familia mantiene su marco de color (capa
   border-box) pero cambia el relleno (padding-box) al degradado de acento,
   para que se vea a la vez "seleccionado" y conserve su color de grupo. */
.nav-btn-fam.active{color:#fff;background:
    linear-gradient(90deg, var(--accent), var(--accent-2)) padding-box,
    linear-gradient(135deg, var(--fam-grad)) border-box;}
/* Un boton activo no debe cambiar de aspecto al pasar el raton por encima.
   Con solo ".active" y ":hover" sueltos, los tres empataban en
   especificidad y quien ganaba dependia del orden de declaracion -- lo que
   hacia que, segun el nivel de menu, el estado "activo" se perdiera visualmente
   al hacer hover (se veia bien en nivel 1 y 3 pero no en nivel 2, por ejemplo).
   Estas dos reglas, con mas especificidad que cualquier ".active" o ":hover"
   suelto, fijan el aspecto activo tambien en hover, en los tres niveles por
   igual. */
.nav-btn.active:hover{background: linear-gradient(90deg, var(--accent), var(--accent-2));color:#fff;border-color: transparent;}
.nav-btn-fam.active:hover{color:#fff;background:
    linear-gradient(90deg, var(--accent), var(--accent-2)) padding-box,
    linear-gradient(135deg, var(--fam-grad)) border-box;}
.main{padding: 18px 16px calc(60px + env(safe-area-inset-bottom, 0));padding-left: calc(16px + env(safe-area-inset-left, 0));padding-right: calc(16px + env(safe-area-inset-right, 0));max-width:1000px;margin:0 auto;}
.panel{display:none;}
.panel.active{display:block; animation: fadein .18s ease;}
@keyframes fadein{from{opacity:0; transform:translateY(4px);} to{opacity:1; transform:translateY(0);}}
.panel-header h2{margin:0 0 4px;font-size: clamp(19px, 5vw, 22px);font-weight:700;}
.panel-toolbar{display:flex;flex-wrap:wrap;gap:12px;align-items:center;justify-content:space-between;margin:16px 0 22px;}
.matrix-hint{color:var(--text-dim); font-size:13px; margin:0;}
.card{background: var(--panel);border:1px solid var(--border);border-radius: var(--radius);margin-bottom:14px;overflow:hidden;}
.card.accessory-card{background: var(--accessory-bg);border:1px dashed var(--accessory-border);}
/* Mismo degradado de familia que el boton del submenu, repetido en el borde
   de la tarjeta para que cabecera y cuerpo desplegable se perciban como el
   mismo grupo del boton que los abre. Al abrir, un lavado sutil (--fam-wash)
   tinta el fondo sin comprometer la legibilidad. */
.card.fam-card{border:1.5px solid transparent;background:
    linear-gradient(var(--panel),var(--panel)) padding-box,
    linear-gradient(135deg, var(--fam-grad)) border-box;}
.card.fam-card.open{background:
    linear-gradient(135deg, var(--fam-wash)) padding-box,
    linear-gradient(var(--panel),var(--panel)) padding-box,
    linear-gradient(135deg, var(--fam-grad)) border-box;}
.card-header{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 16px;min-height:52px;cursor:pointer;user-select:none;touch-action: manipulation;-webkit-tap-highlight-color: transparent;}
/* La cabecera lleva siempre el mismo lavado de color de familia que el boton
   de submenu que la abre (--fam-wash), este abierta o cerrada la tarjeta, para
   que el contraste de color sea visible aunque este colapsada. */
.card.fam-card .card-header{background: linear-gradient(135deg, var(--fam-wash));}
.card-title-wrap{display:flex;align-items:center;gap:10px;min-width:0;}
.chevron{display:inline-block;font-size:11px;color: var(--text-dim);transition: transform .15s ease;flex-shrink:0;}
/* El triangulo de una tarjeta con color de familia usa el mismo color que el
   titulo de la cabecera (no un degradado, que resultaba poco visible). */
.card.fam-card .chevron{color: var(--text);}
.card.open .chevron{ transform: rotate(90deg); }
.card-title{margin:0;font-size:15px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.card-body{max-height:0;overflow:hidden;transition: max-height .25s ease;}
.card.open .card-body{max-height:8000px;}
/* Divisor interno cuando una tarjeta agrupa producto + accesorios (misma
   tarjeta, varias subsecciones separadas por una linea vacia en el Excel). */
.group-divider{padding:10px 16px 6px;border-top:1px solid var(--border);background: var(--accessory-bg);}
.group-divider-text{font-size:12.5px;font-weight:700;color: var(--text-dim);text-transform:uppercase;letter-spacing:.02em;}
/* Resalto de "Accesorios recomendados": solo contorno (sin relleno), en
   rectangulo de cantos redondeados (mismo radio que las tarjetas, no una
   capsula ovalada: asi se lee como una etiqueta dentro del contenido y no
   como un boton/chip de filtro), dibujado con un pseudo-elemento
   enmascarado (mask-composite) igual que el borde con degradado de familia
   de la tarjeta. Se aplica SIEMPRE que la tarjeta tenga color de familia,
   este abierta o cerrada (abrir el producto no abre automaticamente la
   tarjeta de accesorios, asi que el resalto no puede depender de ".open").
   Cubre los TRES sitios donde puede aparecer el texto "Accesorios
   recomendados": el caso mas comun, una fila normal de tabla con fondo
   amarillo dentro de la MISMA tabla de producto (tr.accessory-marker-row,
   sin cabecera ni tarjeta propia); el divisor interno cuando hay dos
   subsecciones en la misma tarjeta (.group-divider-text); y la tarjeta de
   accesorios propia e independiente (titulo .card-title-accessory). */
.card.fam-card .group-divider-text,
.card.fam-card .card-title-accessory,
.card.fam-card tr.accessory-marker-row td:first-child .cell-text{display:inline-flex;align-items:center;position:relative;padding:3px 12px;border-radius:8px;}
.card.fam-card .group-divider-text::before,
.card.fam-card .card-title-accessory::before,
.card.fam-card tr.accessory-marker-row td:first-child .cell-text::before{content:"";position:absolute;inset:0;border-radius:inherit;padding:1.5px;box-sizing:border-box;background: linear-gradient(135deg, var(--fam-grad));-webkit-mask: linear-gradient(#fff 0 0) content-box, linear-gradient(#fff 0 0);-webkit-mask-composite: xor;mask-composite: exclude;pointer-events:none;}
.table-wrap{overflow-x:auto;border-top:1px solid var(--border);}
table{width:100%;border-collapse:collapse;font-size: calc(14.5px * var(--ref-zoom, 1));}
thead th{text-align:left;padding:9px 16px;font-size: calc(12px * var(--ref-zoom, 1));text-transform:uppercase;letter-spacing:.03em;color: var(--text-dim);background: rgba(255,255,255,.02);border-bottom:1px solid var(--border);white-space:nowrap;}
tbody td{padding:9px 16px;border-bottom:1px solid rgba(255,255,255,.04);white-space:nowrap;}
tbody tr:last-child td{border-bottom:none;}
/* Celdas combinadas (notas largas tipo "Requiere un iPhone 11...") no deben
   forzar su frase a una sola linea: eso infla el ancho minimo de toda la
   tabla y dispara el scroll horizontal en movil. Se les permite envolver. */
tbody td[colspan]{white-space:normal;word-break:break-word;}
td.mono, td.mono .cell-text{font-family: var(--mono);font-size: calc(13px * var(--ref-zoom, 1));}
mark{background: #ffdf5e;color:#1a1a1a;padding:0 2px;border-radius:3px;}
.matrix-card{background:var(--panel);}
.matrix-table td{padding:8px 12px;font-size: calc(13px * var(--ref-zoom, 1));border:1px solid rgba(255,255,255,.06);white-space:normal;}
.matrix-table td.th-cell{font-weight:700;}
tr.matrix-spacer td{border:none; padding:4px;}
.no-results{padding:40px 20px;text-align:center;color: var(--text-dim);font-size:14px;}
.search-results-count{color: var(--text-dim);font-size:13.5px;margin:4px 0 18px;}
#searchResultsTable{width:100%;border-collapse:collapse;font-size: calc(14.5px * var(--ref-zoom, 1));}
#searchResultsTable td{padding:10px 14px;border-bottom:1px solid rgba(255,255,255,.06);vertical-align:top;}
#searchResultsTable tr:hover td{background: rgba(255,255,255,.02);}
#searchResultsTable tr.result-accessory td:first-child{border-left:3px solid var(--accessory-border);}
.result-sheet-cell{min-width:180px;white-space:nowrap;}
#searchResultsTable td[colspan] .result-cell-inner,
#searchResultsTable td[colspan] .result-cell-value{white-space:normal;word-break:break-word;}
.result-sheet-link{display:inline-block;font-weight:700;font-size:13.5px;color: var(--accent);cursor:pointer;}
.result-sheet-link:hover{text-decoration: underline;}
.result-section-name{font-size:11.5px;color: var(--text-dim);margin-top:2px;white-space:nowrap;}
.result-cell-inner{display:flex;flex-direction:column;gap:2px;white-space:nowrap;}
.result-cell-label{font-size:10px;text-transform:uppercase;letter-spacing:.03em;color: var(--text-dim);opacity:.75;}
.result-cell-value{font-size: calc(14.5px * var(--ref-zoom, 1));}
#searchResultsTable td.mono .result-cell-value{font-family: var(--mono);font-size: calc(13px * var(--ref-zoom, 1));}
.home-header{margin-bottom:22px;}
.home-header h2{font-size:26px;}
.home-today{display:flex;flex-direction:column;gap:6px;margin:0 0 10px;}
.home-today-date{font-size:13px;font-weight:700;color: var(--accent);text-transform:capitalize;}
.home-tip{color: var(--accent);font-size:13.5px;max-width:640px;margin:0 0 16px;line-height:1.5;}
.home-tip-list{margin:0;padding-left:20px;}
.home-tip-list li{margin:2px 0;}
.home-tip-link{color: var(--accent);font-weight:600;text-decoration: underline dotted var(--accent);text-underline-offset: 3px;}
.home-tip-link:hover{text-decoration-style: solid;}
.home-links{display:flex;flex-direction:column;gap:10px;width:100%;margin:0 0 16px;}
.home-link{display:flex;align-items:center;gap:9px;padding:13px 16px;border:1px solid var(--border);border-radius:12px;background:var(--panel-2);color:var(--text);font-size:14.5px;font-weight:600;text-decoration:none;transition:border-color .15s ease, background .15s ease;box-sizing:border-box;}
.home-link:hover{border-color: var(--accent);background: var(--panel);}
.home-link-icon{font-size:17px;line-height:1;}
.home-link-primary{border-color: var(--accent);}
.help-modal-card{width:460px;}
.help-modal-list{list-style:decimal;margin:0;padding-left:20px;display:flex;flex-direction:column;gap:12px;}
.help-modal-list li{font-size:13.5px;color: var(--text-dim);line-height:1.5;}
.help-modal-list li strong{color: var(--text);font-weight:700;}
.help-inline-icon{vertical-align:middle;margin:0 2px;color: var(--accent);position:relative;top:-1px;}
.incident-modal-card{width:420px;text-align:left;}
.incident-textarea{width:100%;box-sizing:border-box;resize:vertical;min-height:120px;background: var(--panel-2);border:1px solid var(--border);border-radius:10px;padding:10px 12px;font:inherit;font-size:13.5px;color: var(--text);margin-bottom:12px;}
.incident-textarea:focus{outline:none;border-color: var(--accent);}
.incident-actions{display:flex;align-items:center;gap:12px;}
.incident-send-btn{background: var(--accent);color:#fff;border:none;border-radius:8px;padding:9px 18px;font-size:13.5px;font-weight:700;cursor:pointer;transition:opacity .15s ease;}
.incident-send-btn:hover{opacity:.88;}
.incident-send-btn:disabled{opacity:.5;cursor:default;}
.incident-status{font-size:12.5px;color: var(--text-dim);}
.incident-status.ok{color:#3ea25a;}
.incident-status.err{color:#d9534f;}
.home-stats{display:flex;flex-wrap:wrap;gap:22px;font-size:13px;color: var(--text-dim);}
.home-stats strong{color: var(--text);font-size:18px;display:block;}
.home-updated{align-self:center;padding-left:22px;border-left:1px solid var(--border);}
.home-updated strong{color: var(--text);font-weight:600;font-size:13px;display:inline;}
@media (max-width: 880px){
  .main{padding:14px 12px calc(40px + env(safe-area-inset-bottom, 0));}
  .panel-header h2{font-size:19px;}
  .card-title{white-space:normal;}
  /* En pantallas estrechas (no escritorio) empaqueta las columnas a la
     izquierda para no desperdiciar espacio ni forzar scroll horizontal.
     En escritorio (>880px) se mantiene la distribucion simetrica original. */
  .card-body table th:last-child, .card-body table td:last-child{width:100%;}
}
@media (max-width: 480px){
  .topbar-header{padding:8px 12px;}
  .icon-btn{width:36px; height:36px;}
  .btn-row{padding:8px 10px; gap:5px;}
  .btn-row.level2.open{padding:8px 10px;}
  .nav-btn{padding:8px 11px; font-size:12.5px;}
  .nav-btn-leaf{padding:7px 10px; font-size:12px;}
  .main{padding:12px 10px calc(36px + env(safe-area-inset-bottom, 0));}
  .home-stats{gap:16px;}
  .home-updated{padding-left:16px;}
  .panel-header h2{font-size:18px;}
  .card-header{padding:12px 13px; min-height:48px;}
  .card-title{font-size:14px;}
  thead th{padding:7px 8px; font-size: calc(11px * var(--ref-zoom, 1));}
  tbody td{padding:7px 8px; font-size: calc(13.5px * var(--ref-zoom, 1));}
  td.mono, td.mono .cell-text{font-size: calc(12px * var(--ref-zoom, 1));}
  .table-wrap{-webkit-overflow-scrolling:touch;}
}
@media (max-width: 360px){
  .home-stats{flex-direction:column; gap:10px;}
  .home-updated{padding-left:0; border-left:none; border-top:1px solid var(--border); padding-top:10px;}
  thead th{padding:6px 6px;}
  tbody td{padding:6px 6px;}
  td.mono, td.mono .cell-text{font-size: calc(11.5px * var(--ref-zoom, 1));}
}
@supports (padding: max(0px)){
  .btn-row.level1{ padding-left: max(14px, env(safe-area-inset-left)); padding-right: max(14px, env(safe-area-inset-right)); }
}
.ean-clickable{cursor: pointer;text-decoration: underline dotted var(--accent);text-underline-offset: 3px;border-radius: 4px;padding: 1px 4px;transition: background-color .15s ease, color .15s ease;}
.ean-clickable:hover{color: var(--accent) !important;background: rgba(62, 162, 90, 0.15);}
body.light-mode .ean-clickable:hover{background: rgba(11, 122, 62, 0.15);}
.barcode-modal-backdrop,.scanner-modal-backdrop{position: fixed;top: 0;left: 0;width: 100vw;height: 100vh;display: flex;align-items: center;justify-content: center;opacity: 0;pointer-events: none;transition: opacity 0.2s ease;padding: 16px;}
.barcode-modal-backdrop{background: rgba(0, 0, 0, 0.65);backdrop-filter: blur(5px);-webkit-backdrop-filter: blur(5px);z-index: 1000;}
.barcode-modal-backdrop.open{opacity: 1;pointer-events: auto;}
.barcode-modal-card{background: var(--panel);border: 1px solid var(--border);border-radius: 16px;padding: 22px;max-width: 92vw;width: 370px;text-align: center;box-shadow: 0 20px 40px rgba(0, 0, 0, 0.55);transform: translateY(12px) scale(0.96);transition: transform 0.2s cubic-bezier(0.16, 1, 0.3, 1);}
.barcode-modal-backdrop.open .barcode-modal-card{transform: translateY(0) scale(1);}
.barcode-modal-header{display: flex;justify-content: center;align-items: center;margin-bottom: 8px;gap:8px;position:relative;}
.barcode-modal-title{font-size: 13.5px;font-weight: 700;color: var(--text-dim);text-transform: uppercase;letter-spacing: 0.05em;margin: 0;}
/* Por defecto (modo oscuro) la referencia ECI va en un color claro, ya que
   el fondo del marco del modal es oscuro (var(--panel)): un negro fijo
   quedaba ilegible ahi. En modo claro el fondo es claro, asi que se
   mantiene el negro original. */
#barcodeModalTitle{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;text-align:center;color:#f0f2f6;font-size: 20px;}
body.light-mode #barcodeModalTitle{color:#000000;}
#barcodeModalTitle:empty{display:none;}
.barcode-modal-header .barcode-modal-close{position:absolute;right:0;top:50%;transform:translateY(-50%);}
.barcode-modal-header .barcode-share-btn{position:absolute;left:0;top:50%;transform:translateY(-50%);}
.barcode-modal-close,.barcode-share-btn{background: transparent;border: 1px solid var(--border);color: var(--text-dim);font-size: 18px;cursor: pointer;line-height: 1;width: 32px;height: 32px;border-radius: 8px;display: flex;align-items: center;justify-content: center;transition: color 0.15s, border-color 0.15s, background 0.15s;}
.barcode-modal-close:hover,.barcode-share-btn:hover{color: var(--text);border-color: var(--accent);background: var(--panel-2);}
.barcode-share-btn:disabled{opacity:.5;cursor:default;}
.events-modal-card{width: 440px;text-align: left;}
.events-modal-header{display: flex;justify-content: space-between;align-items: center;gap: 12px;margin-bottom: 6px;}
.events-modal-date{font-size: 12.5px;font-weight: 700;color: var(--accent);text-transform: capitalize;margin: 0;}
.events-modal-body{max-height: min(58vh, 420px);overflow-y: auto;-webkit-overflow-scrolling: touch;}
.events-modal-list{list-style: none;padding: 0;margin: 0;display: flex;flex-direction: column;gap: 11px;}
.events-modal-list li{position: relative;padding-left: 17px;font-size: 13.5px;color: var(--text-dim);line-height: 1.5;}
.events-modal-list li:before{content:"\2022";position: absolute;left: 2px;color: var(--accent);font-weight: bold;}
.events-modal-list strong{color: var(--text);font-weight: 700;}
.icon-btn-events{border-color: var(--accent);}
.barcode-product-img-wrap{margin-bottom: 14px;display:flex;align-items:center;justify-content:center;background: var(--panel-2);border: 1px solid var(--border);border-radius: 12px;padding: 10px;}
.barcode-product-img{max-width: 100%;max-height: 220px;object-fit: contain;border-radius: 6px;}
.barcode-svg-container{background: #ffffff;padding: 20px 14px 14px;border-radius: 10px;display: flex;flex-direction: column;align-items: center;justify-content: center;border: 1px solid #dde1e8;}
.barcode-svg-container svg{max-width: 100%;height: auto;display: block;}
.scanner-modal-backdrop{background:rgba(0,0,0,0.75);backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);z-index:1100;}
.scanner-modal-backdrop.open{opacity:1;pointer-events:auto;}
.scanner-modal-card{background:var(--panel);border:1px solid var(--border);border-radius:16px;padding:18px;max-width:94vw;width:420px;max-height:90vh;overflow-y:auto;box-shadow:0 20px 40px rgba(0,0,0,.55);transform:translateY(12px) scale(.96);transition:transform .2s cubic-bezier(.16,1,.3,1);}
.scanner-modal-backdrop.open .scanner-modal-card{transform:translateY(0) scale(1);}
.scanner-modal-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;}
.scanner-modal-title{font-size:13.5px;font-weight:700;color:var(--text-dim);text-transform:uppercase;letter-spacing:.05em;margin:0;}
.scanner-detected{margin-bottom:10px;padding:8px 10px;border-radius:8px;background:var(--panel-2);font-size:12px;word-break:break-all;}
.scanner-detected-label{font-weight:600;color:var(--text-dim);margin-right:6px;}
.scanner-detected-text{font-family:var(--mono);font-size:11.5px;color:var(--text);}
.scanner-status{text-align:center;font-size:13px;color:var(--text-dim);margin-bottom:12px;min-height:20px;}
.scanner-status.found{color:#34c759;font-weight:600;}
.scanner-status.error{color:#ff6b6b;}
.scanner-results{max-height:220px;overflow-y:auto;}
.scanner-no-match{text-align:center;color:var(--text-dim);font-size:13px;padding:16px 8px;}
.scanner-result-item{padding:10px 12px;border:1px solid var(--border);border-radius:9px;margin-bottom:8px;cursor:pointer;transition:border-color .12s,background .12s;touch-action:manipulation;}
.scanner-result-item:hover,.scanner-result-item:active{border-color:var(--accent);background:var(--panel-2);}
.scanner-result-title{font-size:13.5px;font-weight:600;color:var(--text);margin-bottom:4px;}
.scanner-result-detail{display:flex;flex-wrap:wrap;gap:4px;}
.scanner-approx-note{font-size:11px;color:var(--text-dim);margin-bottom:5px;font-style:italic;}
.scanner-detail-chip{font-size:11px;padding:2px 7px;border-radius:6px;background:rgba(255,255,255,.06);color:var(--text-dim);white-space:nowrap;}
.scanner-detail-match{background:rgba(62,162,90,.2);color:var(--accent);font-weight:600;}
body.light-mode .scanner-detail-chip{background:rgba(0,0,0,.05);}
body.light-mode .scanner-detail-match{background:rgba(11,122,62,.12);color:#0b7a3e;}
.scanner-live-wrap{position:relative;border-radius:10px;overflow:hidden;border:1px solid var(--border);background:#000;margin-bottom:12px;display:none;}
.scanner-live-wrap video{width:100%;max-height:280px;object-fit:cover;display:block;opacity:0;transition:opacity .25s ease;}
.scanner-live-wrap.camera-ready video{opacity:1;}
.scanner-live-frame,.scanner-barcode-overlay,.scanner-live-badge,.scanner-live-crop,.scanner-live-overlay{opacity:0;transition:opacity .25s ease;}
.scanner-live-wrap.camera-ready .scanner-live-frame,.scanner-live-wrap.camera-ready .scanner-barcode-overlay,.scanner-live-wrap.camera-ready .scanner-live-badge,.scanner-live-wrap.camera-ready .scanner-live-crop,.scanner-live-wrap.camera-ready .scanner-live-overlay{opacity:1;}
.scanner-live-loading{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;background:#000;color:#fff;font-size:12.5px;font-family:var(--mono);text-align:center;padding:0 16px;transition:opacity .25s ease;z-index:5;}
.scanner-live-wrap.camera-ready .scanner-live-loading{opacity:0;pointer-events:none;}
.scanner-live-spinner{width:26px;height:26px;border-radius:50%;border:3px solid rgba(255,255,255,.25);border-top-color:#fff;animation:scannerSpin .8s linear infinite;margin-bottom:10px;}
@keyframes scannerSpin{to{transform:rotate(360deg);}}
.scanner-live-frame{position:absolute;top:22.5%;left:5%;width:90%;height:55%;border:2px solid rgba(62,162,90,.85);border-radius:10px;pointer-events:none;box-shadow:0 0 0 2000px rgba(0,0,0,.35);transition:top .15s,left .15s,width .15s,height .15s;}
.scanner-live-frame.locked{border-width:3px;border-color:rgba(52,199,89,1);box-shadow:0 0 0 2000px rgba(0,0,0,.55),0 0 14px rgba(52,199,89,.55);}
.scanner-barcode-overlay{position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none;}
.scanner-live-overlay{position:absolute;bottom:0;left:0;right:0;padding:8px 10px;background:linear-gradient(180deg,transparent,rgba(0,0,0,.75));color:#fff;font-size:11.5px;font-family:var(--mono);min-height:18px;}
.scanner-live-badge{position:absolute;top:8px;right:8px;background:rgba(0,0,0,.55);color:#fff;font-size:10.5px;font-weight:700;padding:3px 8px;border-radius:20px;display:flex;align-items:center;gap:5px;letter-spacing:.03em;}
.scanner-live-dot{width:7px;height:7px;border-radius:50%;background:#ff453a;animation:scannerPulse 1.2s ease-in-out infinite;}
@keyframes scannerPulse{0%,100%{opacity:1;}50%{opacity:.3;}}
.scanner-live-stop{display:block;width:100%;padding:9px;border-radius:9px;border:1px solid var(--border);background:var(--panel-2);color:var(--text);font-size:13px;font-weight:600;cursor:pointer;margin-bottom:12px;touch-action:manipulation;}
.scanner-live-stop:active{opacity:.8;}
.scanner-zoom-row{display:none;align-items:center;gap:10px;margin-bottom:12px;padding:0 2px;}
.scanner-zoom-row.visible{display:flex;}
.scanner-zoom-label{font-size:11px;font-weight:700;color:var(--text-dim);text-transform:uppercase;letter-spacing:.04em;white-space:nowrap;}
.scanner-zoom-slider{flex:1;-webkit-appearance:none;appearance:none;height:4px;border-radius:2px;background:var(--panel-2);outline:none;border:1px solid var(--border);}
.scanner-zoom-slider::-webkit-slider-thumb{-webkit-appearance:none;appearance:none;width:22px;height:22px;border-radius:50%;background:var(--accent);cursor:pointer;border:2px solid var(--panel);box-shadow:0 1px 4px rgba(0,0,0,.4);}
.scanner-zoom-slider::-moz-range-thumb{width:22px;height:22px;border-radius:50%;background:var(--accent);cursor:pointer;border:2px solid var(--panel);}
.scanner-zoom-value{font-family:var(--mono);font-size:12px;font-weight:700;color:var(--accent);min-width:38px;text-align:right;}
.scanner-live-crop{position:absolute;top:8px;left:8px;background:rgba(0,0,0,.55);color:#fff;font-size:10px;font-family:var(--mono);padding:3px 7px;border-radius:20px;letter-spacing:.02em;}
"""

TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default" id="statusBarMeta">
<meta name="theme-color" content="#ffffff" id="themeColorMeta">
<title>El Corte Inglés Apple Reference Catalog</title>
<link rel="stylesheet" href="style.css">
</head>
<body class="light-mode">
<div class="topbar">
  <div class="topbar-header">
    <div class="brand-lockup" id="brandHome" role="button" tabindex="0" title="Volver a Inicio" aria-label="Volver a Inicio">
      <svg class="brand-arc-mark" viewBox="0 0 400 228" aria-hidden="true" focusable="false">
        <defs><linearGradient id="eciarcGrad" x1="0%" y1="0%" x2="100%" y2="0%"><stop offset="0%" class="arc-stop-start"/><stop offset="100%" class="arc-stop-end"/></linearGradient></defs>
        <path d="M 30 220 A 170 170 0 0 1 370 220" stroke="url(#eciarcGrad)" stroke-width="11" fill="none"/>
        <path d="M 55 220 A 145 145 0 0 1 345 220" stroke="url(#eciarcGrad)" stroke-width="11" fill="none"/>
        <path d="M 80 220 A 120 120 0 0 1 320 220" stroke="url(#eciarcGrad)" stroke-width="11" fill="none"/>
        <path d="M 105 220 A 95 95 0 0 1 295 220" stroke="url(#eciarcGrad)" stroke-width="11" fill="none"/>
        <path d="M 130 220 A 70 70 0 0 1 270 220" stroke="url(#eciarcGrad)" stroke-width="11" fill="none"/>
        <path d="M 155 220 A 45 45 0 0 1 245 220" stroke="url(#eciarcGrad)" stroke-width="11" fill="none"/>
      </svg>
      <h1 title="El Corte Inglés Apple Reference Catalog"><span class="brand-eci">ECI</span><span class="brand-colon">:</span><span class="brand-arc">arc</span></h1>
    </div>
    <div class="topbar-actions">
      <button id="fontDecBtn" class="icon-btn icon-btn-font" type="button" title="Reducir tamaño de letra" aria-label="Reducir tamaño de letra" style="display:none;">A<span class="font-btn-sign">−</span></button>
      <button id="fontIncBtn" class="icon-btn icon-btn-font" type="button" title="Aumentar tamaño de letra" aria-label="Aumentar tamaño de letra" style="display:none;">A<span class="font-btn-sign">+</span></button>
      <button id="eventsToggle" class="icon-btn icon-btn-events" type="button" title="Un día como hoy en Apple" aria-label="Un día como hoy en Apple" style="display:none;">🐣</button>
      <button id="searchToggle" class="icon-btn" type="button" title="Buscar" aria-label="Buscar">🔎</button>
      <button id="scanToggle" class="icon-btn" type="button" title="Escanear referencia" aria-label="Escanear referencia" style="display:none;">📷</button>
      <button id="themeToggle" class="icon-btn" type="button" title="Cambiar apariencia" aria-label="Cambiar apariencia">🌙</button>
    </div>
  </div>
  <div class="search-bar-row" id="searchBarRow">
    <div class="search-input-wrap">
      <input type="text" id="sheetSearch" placeholder="Buscar por modelo, P/N, ECI o EAN...">
      <button type="button" id="sheetSearchClear" class="search-clear-btn" aria-label="Borrar busqueda" title="Borrar busqueda" style="display:none;">&times;</button>
    </div>
  </div>
  <div class="btn-row level1">__LEVEL1_BTNS__</div>
  __LEVEL2_ROWS__
</div>
<main class="main">
__PANELS__
</main>

<div class="barcode-modal-backdrop" id="barcodeModal" aria-hidden="true" role="dialog">
  <div class="barcode-modal-card">
    <div class="barcode-modal-header">
      <button type="button" class="barcode-share-btn" id="barcodeShareBtn" aria-label="Compartir">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4"/><path d="M8 8l4-4 4 4"/><path d="M4 16v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/></svg>
      </button>
      <h4 class="barcode-modal-title" id="barcodeModalTitle"></h4>
      <button type="button" class="barcode-modal-close" id="barcodeCloseBtn" aria-label="Cerrar">&times;</button>
    </div>
    <div class="barcode-product-img-wrap" id="barcodeProductImgWrap" style="display:none;">
      <img class="barcode-product-img" id="barcodeProductImg" alt="Imagen del producto">
    </div>
    <div class="barcode-svg-container">
      <svg id="barcodeSvg"></svg>
    </div>
  </div>
</div>

<div class="barcode-modal-backdrop events-modal-backdrop" id="eventsModal" aria-hidden="true" role="dialog" aria-label="Un día como hoy en Apple">
  <div class="barcode-modal-card events-modal-card">
    <div class="events-modal-header">
      <p class="events-modal-date" id="eventsModalDate"></p>
      <button type="button" class="barcode-modal-close" id="eventsCloseBtn" aria-label="Cerrar">&times;</button>
    </div>
    <div class="events-modal-body" id="eventsModalBody"></div>
  </div>
</div>

<div class="barcode-modal-backdrop" id="helpModal" aria-hidden="true" role="dialog" aria-labelledby="helpModalTitle">
  <div class="barcode-modal-card events-modal-card help-modal-card">
    <div class="events-modal-header">
      <h4 class="barcode-modal-title" id="helpModalTitle">Cómo usar la página</h4>
      <button type="button" class="barcode-modal-close" id="helpCloseBtn" aria-label="Cerrar">&times;</button>
    </div>
    <div class="events-modal-body">
      <ol class="help-modal-list">
        <li><strong>Menús:</strong> navega entre productos; en los desplegables encontrarás los accesorios relacionados.</li>
        <li><strong>🔎:</strong> busca por EAN, referencia ECI, Part Number o descripción del producto.</li>
        <li><strong>👆 en una línea de producto:</strong> ➡️ ventana con código de barras y foto del artículo. <svg class="help-inline-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 16V4"/><path d="M8 8l4-4 4 4"/><path d="M4 16v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/></svg> para compartir (mensaje, correo, etc.).</li>
        <li><strong>📷:</strong> desde el móvil, apunta a la etiqueta de la caja para identificar el artículo.</li>
        <li>Haz clic sobre el logo <strong>ECI:arc</strong> para volver al inicio y resetear la página.</li>
        <li><strong>Modo claro/oscuro (🌙/☀️):</strong> cambia la apariencia de la página.</li>
        <li><strong>🐣:</strong> solo aparece en fechas señaladas; 👆 para descubrir su contenido.</li>
      </ol>
    </div>
  </div>
</div>

<div class="barcode-modal-backdrop" id="incidentModal" aria-hidden="true" role="dialog" aria-labelledby="incidentModalTitle">
  <div class="barcode-modal-card events-modal-card incident-modal-card">
    <div class="events-modal-header">
      <h4 class="barcode-modal-title" id="incidentModalTitle">Reportar una incidencia</h4>
      <button type="button" class="barcode-modal-close" id="incidentCloseBtn" aria-label="Cerrar">&times;</button>
    </div>
    <textarea id="incidentText" class="incident-textarea" placeholder="Describe lo que quieres reportar..." rows="5"></textarea>
    <div class="incident-actions">
      <button type="button" id="incidentSendBtn" class="incident-send-btn">Enviar</button>
      <span class="incident-status" id="incidentStatus"></span>
    </div>
  </div>
</div>

<div class="scanner-modal-backdrop" id="scannerModal" aria-hidden="true" role="dialog">
  <div class="scanner-modal-card">
    <div class="scanner-modal-header">
      <h4 class="scanner-modal-title">Identificar producto</h4>
      <button type="button" class="barcode-modal-close" id="scannerCloseBtn" aria-label="Cerrar">&times;</button>
    </div>
    <div class="scanner-live-wrap" id="scannerLiveWrap">
      <video id="scannerVideo" playsinline muted autoplay></video>
      <canvas class="scanner-barcode-overlay" id="scannerBarcodeOverlay"></canvas>
      <div class="scanner-live-frame" id="scannerLiveFrame"></div>
      <div class="scanner-live-badge"><span class="scanner-live-dot"></span>EN VIVO</div>
      <div class="scanner-live-crop" id="scannerLiveCrop"></div>
      <div class="scanner-live-overlay" id="scannerLiveOverlay">Iniciando camara...</div>
      <div class="scanner-live-loading"><div><div class="scanner-live-spinner"></div>Preparando base de datos...</div></div>
    </div>
    <div class="scanner-zoom-row" id="scannerZoomRow">
      <span class="scanner-zoom-label">Zoom</span>
      <input type="range" class="scanner-zoom-slider" id="scannerZoomSlider" min="10" max="40" step="1" value="10">
      <span class="scanner-zoom-value" id="scannerZoomValue">1.0x</span>
    </div>
    <button type="button" class="scanner-live-stop" id="scannerLiveStopBtn" style="display:none;">Reiniciar escaner</button>
    <div class="scanner-status" id="scannerStatus"></div>
    <div class="scanner-detected" id="scannerDetected" style="display:none;">
      <span class="scanner-detected-label">Numeros detectados:</span>
      <span class="scanner-detected-text" id="scannerDetectedText"></span>
    </div>
    <div class="scanner-results" id="scannerResults"></div>
  </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/tesseract.js@4/dist/tesseract.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@emailjs/browser@4/dist/email.min.js"></script>
<script>
(function(){
  // --- Load external data (data.json + events.txt) ---
  var ALL_ROWS = [];
  var APPLE_ON_THIS_DAY = {};
  // {hoja: [ean, ...]} de las fotos .webp presentes en disco, generado en el
  // build. Se consulta antes de pedir una imagen para no lanzar peticiones 404.
  var IMAGE_MANIFEST = {};

  function initApp() {
    var panels = document.querySelectorAll('.panel');
    var sheetSearch = document.getElementById('sheetSearch');
    var sheetSearchClear = document.getElementById('sheetSearchClear');
    var level1Buttons = document.querySelectorAll('.btn-row.level1 .nav-btn');
    var lobRows = document.querySelectorAll('.btn-row.level2');
    var leafButtons = document.querySelectorAll('.nav-btn-leaf');

    // --- Show today's ephemeris (in a modal, opened from the topbar gift icon) ---
    var homeTodayDate = document.getElementById('homeTodayDate');
    var eventsToggle = document.getElementById('eventsToggle');
    var eventsModal = document.getElementById('eventsModal');
    var eventsCloseBtn = document.getElementById('eventsCloseBtn');
    var eventsModalDate = document.getElementById('eventsModalDate');
    var eventsModalBody = document.getElementById('eventsModalBody');

    function escHtml(s){
      return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    // Los tres modales (eventos, codigo de barras, escaner) comparten el mismo
    // interruptor abierto/cerrado via clase + aria-hidden.
    function setModalOpen(modalEl, open){
      if(!modalEl) return;
      modalEl.classList.toggle('open', open);
      modalEl.setAttribute('aria-hidden', open ? 'false' : 'true');
    }

    function openEventsModal(){
      setModalOpen(eventsModal, true);
    }
    function closeEventsModal(){
      setModalOpen(eventsModal, false);
    }

    var today = new Date();
    var weekday = today.toLocaleDateString('es-ES', { weekday: 'long' });
    weekday = weekday.charAt(0).toUpperCase() + weekday.slice(1);
    var dd = String(today.getDate()).padStart(2, '0');
    var mm = String(today.getMonth() + 1).padStart(2, '0');
    var yyyy = today.getFullYear();
    var todayLabel = weekday + ' ' + dd + '/' + mm + '/' + yyyy;
    if(homeTodayDate){ homeTodayDate.textContent = todayLabel; }

    var todayEvents = APPLE_ON_THIS_DAY[mm + '-' + dd];
    if(todayEvents && todayEvents.length){
      if(eventsModalDate){ eventsModalDate.textContent = todayLabel; }
      if(eventsModalBody){
        var items = todayEvents.map(function(ev){
          return '<li><strong>' + dd + '/' + mm + '/' + ev.y + '</strong>: ' + escHtml(ev.t) + '</li>';
        }).join('');
        eventsModalBody.innerHTML = '<ul class="events-modal-list">' + items + '</ul>';
      }
      if(eventsToggle){
        eventsToggle.style.display = '';
        eventsToggle.addEventListener('click', function(){
          if(eventsModal && eventsModal.classList.contains('open')) closeEventsModal();
          else openEventsModal();
        });
      }
    } else if(eventsToggle){
      // No ephemeris today: the extra button simply never shows up.
      eventsToggle.style.display = 'none';
    }

    if(eventsCloseBtn){ eventsCloseBtn.addEventListener('click', closeEventsModal); }
    if(eventsModal){
      eventsModal.addEventListener('click', function(e){
        if(e.target === eventsModal) closeEventsModal();
      });
    }
    document.addEventListener('keydown', function(e){
      if(e.key === 'Escape' && eventsModal && eventsModal.classList.contains('open')){
        closeEventsModal();
      }
    });

    // --- Modal de Ayuda (se abre desde el enlace del landing) ---
    var homeHelpLink = document.getElementById('homeHelpLink');
    var helpModal = document.getElementById('helpModal');
    var helpCloseBtn = document.getElementById('helpCloseBtn');
    function openHelpModal(){ setModalOpen(helpModal, true); }
    function closeHelpModal(){ setModalOpen(helpModal, false); }
    if(homeHelpLink){
      homeHelpLink.addEventListener('click', function(e){
        e.preventDefault();
        openHelpModal();
      });
    }
    if(helpCloseBtn){ helpCloseBtn.addEventListener('click', closeHelpModal); }
    if(helpModal){
      helpModal.addEventListener('click', function(e){
        if(e.target === helpModal) closeHelpModal();
      });
    }
    document.addEventListener('keydown', function(e){
      if(e.key === 'Escape' && helpModal && helpModal.classList.contains('open')){
        closeHelpModal();
      }
    });

    // --- Modal de incidencia: envia el texto por EmailJS, sin salir de la pagina ---
    var EMAILJS_PUBLIC_KEY = 'LUJwdk3MQpdADb5Nm';
    var EMAILJS_SERVICE_ID = 'service_jqp3osl';
    var EMAILJS_TEMPLATE_ID = 'template_1146g92';
    if(window.emailjs && typeof window.emailjs.init === 'function'){
      window.emailjs.init({ publicKey: EMAILJS_PUBLIC_KEY });
    }
    var homeIncidentLink = document.getElementById('homeIncidentLink');
    var incidentModal = document.getElementById('incidentModal');
    var incidentCloseBtn = document.getElementById('incidentCloseBtn');
    var incidentText = document.getElementById('incidentText');
    var incidentSendBtn = document.getElementById('incidentSendBtn');
    var incidentStatus = document.getElementById('incidentStatus');
    function openIncidentModal(){
      if(incidentStatus){ incidentStatus.textContent = ''; incidentStatus.className = 'incident-status'; }
      if(incidentSendBtn){ incidentSendBtn.disabled = false; incidentSendBtn.textContent = 'Enviar'; }
      setModalOpen(incidentModal, true);
      if(incidentText){ setTimeout(function(){ incidentText.focus(); }, 50); }
    }
    function closeIncidentModal(){
      setModalOpen(incidentModal, false);
      if(incidentText){ incidentText.value = ''; }
    }
    if(homeIncidentLink){
      homeIncidentLink.addEventListener('click', function(e){
        e.preventDefault();
        openIncidentModal();
      });
    }
    if(incidentCloseBtn){ incidentCloseBtn.addEventListener('click', closeIncidentModal); }
    if(incidentModal){
      incidentModal.addEventListener('click', function(e){
        if(e.target === incidentModal) closeIncidentModal();
      });
    }
    document.addEventListener('keydown', function(e){
      if(e.key === 'Escape' && incidentModal && incidentModal.classList.contains('open')){
        closeIncidentModal();
      }
    });
    if(incidentSendBtn){
      incidentSendBtn.addEventListener('click', function(){
        var msg = incidentText ? incidentText.value.trim() : '';
        if(!msg){
          if(incidentStatus){ incidentStatus.textContent = 'Escribe algo antes de enviar.'; incidentStatus.className = 'incident-status err'; }
          return;
        }
        if(!window.emailjs){
          if(incidentStatus){ incidentStatus.textContent = 'No se pudo cargar el servicio de envio.'; incidentStatus.className = 'incident-status err'; }
          return;
        }
        incidentSendBtn.disabled = true;
        incidentSendBtn.textContent = 'Enviando...';
        if(incidentStatus){ incidentStatus.textContent = ''; incidentStatus.className = 'incident-status'; }
        window.emailjs.send(EMAILJS_SERVICE_ID, EMAILJS_TEMPLATE_ID, { message: msg }).then(function(){
          if(incidentStatus){ incidentStatus.textContent = 'Enviado. ¡Gracias!'; incidentStatus.className = 'incident-status ok'; }
          incidentSendBtn.textContent = 'Enviado';
          if(incidentText){ incidentText.value = ''; }
          setTimeout(closeIncidentModal, 1400);
        }).catch(function(){
          if(incidentStatus){ incidentStatus.textContent = 'No se pudo enviar. Intentalo de nuevo.'; incidentStatus.className = 'incident-status err'; }
          incidentSendBtn.disabled = false;
          incidentSendBtn.textContent = 'Enviar';
        });
      });
    }

    // --- Theme: claro por defecto; el usuario puede pasar a oscuro con la luna ---
    var themeToggle = document.getElementById('themeToggle');
    var statusBarMeta = document.getElementById('statusBarMeta');
    var themeColorMeta = document.getElementById('themeColorMeta');
    function applyTheme(theme){
      document.body.classList.toggle('light-mode', theme === 'light');
      if(themeToggle){ themeToggle.textContent = theme === 'light' ? '🌙' : '☀️'; }
      if(statusBarMeta){ statusBarMeta.setAttribute('content', theme === 'light' ? 'default' : 'black'); }
      if(themeColorMeta){ themeColorMeta.setAttribute('content', theme === 'light' ? '#ffffff' : '#131620'); }
    }

    var THEME_KEY = 'catalogTheme';
    var savedTheme = null;
    try{ savedTheme = localStorage.getItem(THEME_KEY); }catch(e){}
    applyTheme(savedTheme === 'dark' ? 'dark' : 'light');

    if(themeToggle){
      themeToggle.addEventListener('click', function(){
        var isLight = document.body.classList.contains('light-mode');
        var next = isLight ? 'dark' : 'light';
        applyTheme(next);
        try{ localStorage.setItem(THEME_KEY, next); }catch(e){}
      });
    }

    // --- Tamaño de letra (A- / A+) ---
    // La variable --ref-zoom multiplica (via calc()) el font-size real de las
    // celdas de las tablas de referencia (secciones de producto, matriz de
    // compatibilidad y resultados de busqueda) — no se usa la propiedad CSS
    // "zoom" porque Safari/iOS no la soporta. Asi escala solo el contenido de
    // los campos de referencia y no afecta a cabeceras, botones ni al resto
    // de la pagina. Se guarda en localStorage para que persista entre visitas.
    var FONT_SCALE_KEY = 'fontScale';
    var FONT_SCALE_MIN = 0.8, FONT_SCALE_MAX = 1.4, FONT_SCALE_STEP = 0.1;
    var fontDecBtn = document.getElementById('fontDecBtn');
    var fontIncBtn = document.getElementById('fontIncBtn');
    var fontScale = 1;
    try{
      var storedScale = parseFloat(localStorage.getItem(FONT_SCALE_KEY));
      if(!isNaN(storedScale)){ fontScale = storedScale; }
    }catch(e){}

    function applyFontScale(){
      document.documentElement.style.setProperty('--ref-zoom', fontScale);
      if(fontIncBtn){ fontIncBtn.disabled = fontScale >= FONT_SCALE_MAX - 1e-9; }
      if(fontDecBtn){ fontDecBtn.disabled = fontScale <= FONT_SCALE_MIN + 1e-9; }
    }
    function setFontScale(next){
      fontScale = Math.min(FONT_SCALE_MAX, Math.max(FONT_SCALE_MIN, Math.round(next * 100) / 100));
      try{ localStorage.setItem(FONT_SCALE_KEY, fontScale); }catch(e){}
      applyFontScale();
    }
    applyFontScale();
    if(fontIncBtn){ fontIncBtn.addEventListener('click', function(){ setFontScale(fontScale + FONT_SCALE_STEP); }); }
    if(fontDecBtn){ fontDecBtn.addEventListener('click', function(){ setFontScale(fontScale - FONT_SCALE_STEP); }); }

    // Los botones A-/A+ solo tienen sentido cuando hay una tabla de referencias
    // en pantalla (hoja de producto o resultados de busqueda); se ocultan en Inicio.
    function toggleFontButtons(show){
      var disp = show ? 'flex' : 'none';
      if(fontIncBtn){ fontIncBtn.style.display = disp; }
      if(fontDecBtn){ fontDecBtn.style.display = disp; }
    }

    // --- Search bar toggle ---
    var searchToggle = document.getElementById('searchToggle');
    var searchBarRow = document.getElementById('searchBarRow');
    searchToggle.addEventListener('click', function(){
      var open = searchBarRow.classList.toggle('open');
      if(open){ sheetSearch.focus(); }
    });

    // --- Navigation logic ---
    function setActiveLevel1(el){
      level1Buttons.forEach(function(b){ b.classList.remove('active'); });
      if(el){ el.classList.add('active'); }
    }
    function setActiveLeaf(slug){
      leafButtons.forEach(function(b){ b.classList.toggle('active', b.dataset.target === slug); });
    }
    function hideAllLobRows(){
      lobRows.forEach(function(r){ r.classList.remove('open'); });
    }
    function showLobRow(lobSlug){
      lobRows.forEach(function(r){ r.classList.toggle('open', r.id === 'lobrow-' + lobSlug); });
    }
    function findLobForSlug(slug){
      var leaf = document.querySelector('.nav-btn-leaf[data-target="' + slug + '"]');
      if(leaf){
        var row = leaf.closest('.btn-row.level2');
        if(row){ return row.id.replace('lobrow-', ''); }
      }
      return null;
    }
    function syncNav(slug){
      if(slug === 'home' || slug === 'search-results'){
        // null: no hay boton de nivel 1 que representar como activo en Inicio ni
        // en resultados de busqueda, y setActiveLevel1(null) limpia el resto.
        setActiveLevel1(null);
        hideAllLobRows();
        setActiveLeaf(null);
        return;
      }
      var directBtn = document.querySelector('.btn-row.level1 .nav-btn[data-target="' + slug + '"]');
      if(directBtn){
        setActiveLevel1(directBtn);
        hideAllLobRows();
        setActiveLeaf(null);
        return;
      }
      var lobSlug = findLobForSlug(slug);
      if(lobSlug){
        var lobBtn = document.querySelector('.btn-row.level1 .nav-btn[data-lob="' + lobSlug + '"]');
        setActiveLevel1(lobBtn);
        showLobRow(lobSlug);
        setActiveLeaf(slug);
      }
    }

    // --- Reset to Home ---
    function resetToHome(opts){
      var keepSearchOpen = !!(opts && opts.keepSearchOpen);
      panels.forEach(function(p){ p.classList.toggle('active', p.id === 'home'); });
      level1Buttons.forEach(function(b){ b.classList.remove('active'); });
      lobRows.forEach(function(r){ r.classList.remove('open'); });
      leafButtons.forEach(function(b){ b.classList.remove('active'); });
      collapseAllCards();
      if(!keepSearchOpen){
        if(searchDebounceTimer){ clearTimeout(searchDebounceTimer); searchDebounceTimer = null; }
        if(sheetSearch){ sheetSearch.value = ''; sheetSearch.blur(); }
        if(sheetSearchClear){ sheetSearchClear.style.display = 'none'; }
        if(searchBarRow){ searchBarRow.classList.remove('open'); }
      }
      if(resultsBody){ resultsBody.innerHTML = ''; }
      if(resultsCount){ resultsCount.textContent = ''; }
      toggleFontButtons(false);
      var existingEmpty = document.querySelector('#search-results .no-results');
      if(existingEmpty){ existingEmpty.remove(); }
      if(history.replaceState){ history.replaceState(null, '', '#home'); }
      window.scrollTo({top: 0, behavior: 'smooth'});
    }

    function activate(slug, opts){
      if(slug === 'home'){ resetToHome(opts); return; }
      collapseAllCards();
      panels.forEach(function(p){ p.classList.toggle('active', p.id === slug); });
      syncNav(slug);
      toggleFontButtons(true);
      if(history.replaceState){ history.replaceState(null, '', '#' + slug); }
    }

    // Al navegar a otro LOB/hoja/submenu se contrae cualquier tarjeta de
    // producto que hubiera quedado desplegada, para que cada seccion empiece
    // siempre colapsada.
    function collapseAllCards(){
      document.querySelectorAll('.card.open').forEach(function(c){ c.classList.remove('open'); });
      document.querySelectorAll('.nav-btn-model.active').forEach(function(b){ b.classList.remove('active'); });
    }

    function deactivateAllPanels(){
      panels.forEach(function(p){ p.classList.remove('active'); });
    }

    level1Buttons.forEach(function(btn){
      btn.addEventListener('click', function(e){
        e.preventDefault();
        if(btn.dataset.lob){
          var lobSlug = btn.dataset.lob;
          collapseAllCards();
          setActiveLevel1(btn);
          showLobRow(lobSlug);
          setActiveLeaf(null);
          deactivateAllPanels();
          if(history.replaceState){ history.replaceState(null, '', '#'); }
        } else if(btn.dataset.target){
          hideAllLobRows();
          activate(btn.dataset.target);
          window.scrollTo({top:0, behavior:'smooth'});
        }
      });
    });

    // La zona de marca (arcos + ECIarc) hace de boton Inicio: clic o Enter/Espacio.
    var brandHome = document.getElementById('brandHome');
    if(brandHome){
      brandHome.addEventListener('click', function(e){
        e.preventDefault();
        resetToHome();
      });
      brandHome.addEventListener('keydown', function(e){
        if(e.key === 'Enter' || e.key === ' ' || e.key === 'Spacebar'){
          e.preventDefault();
          resetToHome();
        }
      });
    }

    leafButtons.forEach(function(btn){
      btn.addEventListener('click', function(e){
        e.preventDefault();
        if(btn.dataset.target){
          activate(btn.dataset.target);
          window.scrollTo({top:0, behavior:'smooth'});
        }
      });
    });

    // Abre/cierra una tarjeta y, si tiene una tarjeta de accesorios (o de
    // producto) emparejada via data-pair-card, la sincroniza al mismo estado.
    // Marca como "activo" el boton de nivel 3 (si lo hay) que apunta a esta
    // tarjeta, para que el resalto de color siga el estado real de apertura.
    function syncModelButtonActive(card){
      if(!card) return;
      var btn = document.querySelector('.nav-btn-model[data-model-target="' + card.id + '"]');
      if(btn){ btn.classList.toggle('active', card.classList.contains('open')); }
    }

    // Abre/cierra una tarjeta. El cierre SI se propaga a su tarjeta emparejada
    // (producto <-> accesorios), pero la apertura NO: desplegar los accesorios
    // no debe desplegar automaticamente el articulo, y viceversa.
    function setCardOpen(card, open){
      if(!card) return;
      card.classList.toggle('open', open);
      syncModelButtonActive(card);
      if(!open){
        var pairId = card.dataset.pairCard;
        if(pairId){
          var partner = document.getElementById(pairId);
          if(partner){
            partner.classList.remove('open');
            syncModelButtonActive(partner);
          }
        }
      }
    }

    // Level-3 model buttons
    document.querySelectorAll('.nav-btn-model').forEach(function(btn){
      btn.addEventListener('click', function(e){
        e.preventDefault();
        var targetId = btn.dataset.modelTarget;
        var card = document.getElementById(targetId);
        if(!card){ return; }
        if(card.classList.contains('open')){
          setCardOpen(card, false);
          return;
        }
        setCardOpen(card, true);
        var topbar = document.querySelector('.topbar');
        var headerOffset = topbar ? topbar.offsetHeight + 12 : 60;
        var top = card.getBoundingClientRect().top + window.pageYOffset - headerOffset;
        window.scrollTo({top: Math.max(0, top), behavior:'smooth'});
      });
    });

    // Card collapse / expand
    document.querySelectorAll('[data-toggle]').forEach(function(header){
      header.addEventListener('click', function(){
        var card = header.closest('.card');
        if(card){ setCardOpen(card, !card.classList.contains('open')); }
      });
    });

    // Global search

    // Diccionario ES -> EN generado en el build a partir del fichero "Colors"
    // del proyecto. Clave = lo que teclea el empleado (espanol, ya normalizado
    // sin acentos ni marca de plural/genero); valor = terminos tal como
    // aparecen en la base de datos (ingles). Para actualizarlo se edita Colors
    // y se relanza el build: no se toca este fichero.
    var TERM_SYNONYMS = __TERM_SYNONYMS__;
    var TERM_MAX_WORDS = __TERM_MAX_WORDS__;
    var TERM_KEYS = Object.keys(TERM_SYNONYMS);
    // Longitud minima de un prefijo para disparar la busqueda incremental en
    // el diccionario (por debajo de esto hay demasiadas claves candidatas y
    // el resultado deja de ser util).
    var TERM_PREFIX_MIN = 2;

    // Palabras de relleno del lenguaje natural (preposiciones, articulos,
    // conjunciones). Si el empleado escribe "reloj de titanio dorado con correa
    // corta", los conectores "de"/"con" no son datos: se descartan para llegar
    // al resultado. Solo se descartan cuando quedan SUELTOS entre terminos; si
    // forman parte de una traduccion del diccionario ("correa de muñeca",
    // "todo en uno") el emparejado por frase los absorbe antes y no llegan aqui.
    // Se comparan contra la forma sin acentos del token.
    var STOPWORDS = {
      'de':1,'del':1,'con':1,'y':1,'o':1,'u':1,'e':1,'para':1,'por':1,
      'el':1,'la':1,'los':1,'las':1,'un':1,'una':1,'unos':1,'unas':1,
      'en':1,'al':1,'a':1,'sin':1,'su':1,'sus':1,'lo':1,
      'the':1,'of':1,'with':1,'and':1,'for':1
    };

    // Minusculas y sin acentos, SIN tocar plural ni genero. Es el nivel con el
    // que se casa el texto literal (P/N, EAN, ECI, descripcion) y los terminos
    // del diccionario contra la fila: conserva la grafia real del dato.
    function stripAccents(s){
      return String(s).toLowerCase()
        .replace(/[áàäâã]/g,'a').replace(/[éèëê]/g,'e')
        .replace(/[íìïî]/g,'i').replace(/[óòöôõ]/g,'o')
        .replace(/[úùüû]/g,'u').replace(/ñ/g,'n').replace(/ç/g,'c');
    }

    // stripAccents() y ademas sin marca de plural ni genero, para que
    // "Pequeñas", "pequeno" y "PEQUEÑA" acaben en la misma clave ("pequeno").
    // Tiene que producir lo mismo que fold_es() en build_html.py.
    // Solo recorta palabras puramente alfabeticas: un P/N ("MHFA4Y/A"), un EAN,
    // una ECI o una talla ("S/M") llevan digitos o barras y quedan intactos.
    // Este nivel se usa SOLO para buscar la clave en el diccionario, nunca para
    // casar contra la fila.
    function foldES(s){
      var words = stripAccents(s).split(/\s+/), out = [];
      for(var i=0;i<words.length;i++){
        var w = words[i];
        if(!w){ continue; }
        if(/^[a-z]+$/.test(w)){
          if(w.length > 3 && w.charAt(w.length-1) === 's'){ w = w.slice(0,-1); }
          if(w.length > 3 && w.charAt(w.length-1) === 'a'){ w = w.slice(0,-1) + 'o'; }
        }
        out.push(w);
      }
      return out.join(' ');
    }

    // Los equivalentes ingleses se comparan por palabra completa, no por
    // substring: valores cortos del diccionario como "S", "M" o "S/M"
    // apareceran dentro de casi cualquier fila si se buscan como substring.
    var _altReCache = {};
    function altMatches(hay, alt){
      var re = _altReCache[alt];
      if(!re){
        re = new RegExp('(^|[^a-z0-9])' + escapeRegExp(alt) + '([^a-z0-9]|$)');
        _altReCache[alt] = re;
      }
      return re.test(hay);
    }

    // Convierte lo tecleado en grupos de condiciones. Cada grupo debe cumplirse
    // (AND entre grupos, OR entre las alternativas de un grupo). Se reconocen
    // traducciones de varias palabras ("blanco estrella") buscando primero la
    // coincidencia mas larga.
    //
    // Resultados en tiempo real sin terminar de escribir: si no hay clave EXACTA
    // para el trozo tecleado, se busca por PREFIJO entre las claves del
    // diccionario ("estr" encuentra la clave "estrella" aunque falten letras).
    // Solo el ULTIMO grupo (lo que el empleado esta escribiendo en este
    // instante) se trata como posiblemente incompleto; los anteriores, al ya
    // tener un espacio detras, se tratan como palabras cerradas. Si el prefijo
    // es ambiguo (casa con varias claves: "pl" -> plata, plastico...) se
    // combinan todas las traducciones candidatas.
    function buildGroups(raw){
      var endsWithSpace = /\s$/.test(raw);
      var lit = raw.toLowerCase().split(/\s+/).filter(Boolean);
      var acc = stripAccents(raw).split(/\s+/).filter(Boolean);
      var fold = foldES(raw).split(/\s+/).filter(Boolean);
      var groups = [], i = 0;
      while(i < fold.length){
        var hit = null, end = i + 1;
        var maxJ = Math.min(fold.length, i + TERM_MAX_WORDS);
        for(var j = maxJ; j > i; j--){
          var phrase = fold.slice(i, j).join(' ');
          var vals = TERM_SYNONYMS[phrase];
          if(vals){ hit = vals; end = j; break; }
          // Prefijo: solo si este trozo es lo ultimo que hay tecleado (todavia
          // puede faltar por escribir) y tiene longitud suficiente. Se compara
          // el prefijo en DOS formas, porque el plegado de genero/plural (a->o,
          // -s) altera la ultima letra de la palabra a medio escribir y rompe
          // la coincidencia ("media" plegado seria "medio", que no es prefijo
          // de "medianoche"): (1) plegado, para tolerar genero/plural ya
          // completos ("rojas"->"rojo"); (2) solo sin acentos, que conserva la
          // palabra parcial intacta ("media" sigue siendo "media").
          var phraseAcc = acc.slice(i, j).join(' ');
          var atTail = (j === fold.length) && !endsWithSpace;
          if(atTail && phraseAcc.length >= TERM_PREFIX_MIN){
            var merged = null;
            for(var k = 0; k < TERM_KEYS.length; k++){
              var tk = TERM_KEYS[k];
              if(tk.indexOf(phrase) === 0 || tk.indexOf(phraseAcc) === 0){
                merged = merged || [];
                var cand = TERM_SYNONYMS[tk];
                for(var m = 0; m < cand.length; m++){
                  if(merged.indexOf(cand[m]) === -1){ merged.push(cand[m]); }
                }
              }
            }
            if(merged){ hit = merged; end = j; break; }
          }
        }
        // Stopwords del lenguaje natural: si una palabra suelta (no absorbida por
        // ninguna frase del diccionario ni traducida) es un conector ("de", "con"
        // ...), se descarta para que "reloj de titanio ... con correa corta"
        // llegue al resultado. No se descarta si es el ultimo token aun tecleado
        // (podria ser prefijo de un dato real: "de"->"desierto"). Se compara con
        // la forma SIN acentos (acc), porque el plegado deforma los conectores
        // ("para"->"paro"). Las traducciones multi-palabra con conectores
        // ("correa de muñeca", "todo en uno") ya han sido absorbidas arriba por
        // el emparejado por frase (greedy, de mas largo a mas corto).
        if(hit === null && end === i + 1){
          var sw = acc.slice(i, end).join(' ');
          var tail = (end === fold.length) && !endsWithSpace;
          if(!tail && STOPWORDS[sw]){ i = end; continue; }
        }
        groups.push({
          lit: lit.slice(i, end).join(' '),
          acc: acc.slice(i, end).join(' '),
          fold: fold.slice(i, end).join(' '),
          alts: hit
        });
        i = end;
      }
      return groups;
    }

    function escapeRegExp(s){ return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
    // Compila una regex de resaltado UNA vez por busqueda. Antes se reconstruia
    // dentro de highlightText() en cada celda renderizada (3 llamadas por fila
    // mostrada), repitiendo la misma compilacion cientos de veces por tecla
    // pulsada sobre resultados que no cambian de token dentro de la pasada.
    function buildHighlightRegex(tokens){
      if(!tokens || !tokens.length){ return null; }
      var sorted = tokens.slice().sort(function(a, b){ return b.length - a.length; });
      var pattern = sorted.map(escapeRegExp).join('|');
      return pattern ? new RegExp('(' + pattern + ')', 'ig') : null;
    }
    function highlightText(text, re){
      if(!re){ return text; }
      return String(text).replace(re, '<mark>$1</mark>');
    }

    var resultsBody = document.getElementById('searchResultsBody');
    var resultsCount = document.getElementById('searchResultsCount');

    function renderResults(term, tokens){
      resultsBody.innerHTML = '';
      var groups = buildGroups(term);
      // Para resaltar se usa lo tecleado mas los equivalentes ingleses, de modo
      // que al buscar "burdeos" se marque el "Burgundy" que trae la fila.
      var hlTokens = [];
      groups.forEach(function(g){
        if(g.lit && hlTokens.indexOf(g.lit) === -1){ hlTokens.push(g.lit); }
        if(g.alts){
          g.alts.forEach(function(a){ if(hlTokens.indexOf(a) === -1){ hlTokens.push(a); } });
        }
      });
      var hlRe = buildHighlightRegex(hlTokens);
      var matches = ALL_ROWS.filter(function(r){
        // r._sa es r.search sin acentos, calculado una sola vez al cargar. Se
        // casa SIEMPRE contra este texto, que conserva la grafia real del dato
        // (P/N, EAN, ECI, descripcion); el plegado de plural/genero solo se usa
        // para encontrar la clave del diccionario.
        var hay = r._sa || (r._sa = stripAccents(r.search));
        return groups.every(function(g){
          // 1) Lo tecleado, tal cual, como substring: permite ir escribiendo a
          //    medias ("macb", "mhfa") y cubre P/N, descripcion e ingles directo.
          if (hay.indexOf(g.acc) !== -1) return true;
          // 1b) Igual pero sobre el texto plegado (plural/genero): asi "indigos"
          //    o "indigos" encuentran el dato ingles "Indigo" directamente, y en
          //    general la busqueda literal tolera plural y genero ("correas" ->
          //    "Correa"). r._sf es r.search plegado, calculado una sola vez.
          if (g.fold) {
            var hayF = r._sf || (r._sf = foldES(r.search));
            if (hayF.indexOf(g.fold) !== -1) return true;
          }
          // 2) Equivalentes del diccionario, por palabra completa (si se
          //    buscasen como substring, valores cortos como "S" o "M" saldrian
          //    en casi cualquier fila).
          if (g.alts) {
            for (var i = 0; i < g.alts.length; i++) {
              if (altMatches(hay, stripAccents(g.alts[i]))) return true;
            }
          }
          // Si el termino es una referencia numerica (solo digitos y
          // separadores: puntos, espacios, guiones), se compara tambien su
          // forma solo-digitos contra la de la fila. Asi "152.170.12424",
          // "15217012424" o "152.1701.2424" encuentran lo mismo: lo que importa
          // son los 11 digitos, no como esten partidos en la etiqueta. Se exige
          // longitud >= 6 para no convertir un "24" suelto en coincidencias.
          if (/^[\d.\-\s·,/]+$/.test(g.lit)) {
            var td = g.lit.replace(/\D/g, '');
            if (td.length >= 6 && r.searchnum.indexOf(td) !== -1) return true;
          }
          return false;
        });
      });
      resultsCount.textContent = matches.length + ' referencia' + (matches.length === 1 ? '' : 's') + ' encontrada' + (matches.length === 1 ? '' : 's') + ' para "' + term + '"';

      var existingEmpty = document.querySelector('#search-results .no-results');
      if(existingEmpty){ existingEmpty.remove(); }

      if(matches.length === 0){
        var emptyRow = document.createElement('div');
        emptyRow.className = 'no-results';
        emptyRow.textContent = 'Sin resultados para "' + term + '"';
        resultsBody.parentElement.parentElement.appendChild(emptyRow);
        return;
      }

      matches.forEach(function(r){
        var tr = document.createElement('tr');
        if(r.isAccessory){ tr.classList.add('result-accessory'); }
        tr.dataset.sheetFolder = r.sheetFolder || '';
        var eanCell = r.cells.filter(function(c){ return c.label && c.label.trim().toUpperCase() === 'EAN'; })[0];
        var rowEan = (eanCell && eanCell.v && /^\d{8,14}$/.test(eanCell.v.trim())) ? eanCell.v.trim() : '';
        tr.dataset.ean = rowEan;
        var eciCell = r.cells.filter(function(c){ return c.label && c.label.trim().toUpperCase() === 'ECI'; })[0];
        tr.dataset.eci = (eciCell && eciCell.v) ? eciCell.v.trim() : '';

        var tdSheet = document.createElement('td');
        tdSheet.className = 'result-sheet-cell';
        var link = document.createElement('a');
        link.href = '#' + r.slug;
        link.className = 'result-sheet-link';
        link.innerHTML = highlightText(r.sheet, hlRe);
        var secSpan = document.createElement('div');
        secSpan.className = 'result-section-name';
        secSpan.innerHTML = (r.isAccessory ? '🔌 ' : '') + highlightText(r.section, hlRe);
        link.addEventListener('click', function(e){
          e.preventDefault();
          activate(r.slug);
          window.scrollTo({top:0, behavior:'smooth'});
        });
        tdSheet.appendChild(link);
        tdSheet.appendChild(secSpan);
        tr.appendChild(tdSheet);

        r.cells.forEach(function(c){
          var td = document.createElement('td');
          if(c.span){ td.colSpan = c.span; }
          if(c.mono){ td.classList.add('mono'); }
          if(c.bg){
            td.style.background = c.bg;
            td.style.color = c.fg || '#111318';
          }
          var wrap = document.createElement('div');
          wrap.className = 'result-cell-inner';
          var labelSpan = document.createElement('span');
          labelSpan.className = 'result-cell-label';
          labelSpan.textContent = c.label;
          var valueSpan = document.createElement('span');
          valueSpan.className = 'result-cell-value';
          valueSpan.innerHTML = c.v ? highlightText(c.v, hlRe) : '';
          if (rowEan) { valueSpan.classList.add('ean-clickable'); valueSpan.title = 'Ver código de barras'; }
          wrap.appendChild(labelSpan);
          wrap.appendChild(valueSpan);
          td.appendChild(wrap);
          tr.appendChild(td);
        });

        resultsBody.appendChild(tr);
      });
    }

    // Debounce de la busqueda en vivo: sin el, cada tecla pulsada disparaba un
    // Array.filter sobre ~2000 filas mas la reconstruccion completa del DOM de
    // resultados (renderResults), aunque la tecla siguiente llegara 50ms
    // despues y desechara ese trabajo de inmediato. 130ms es imperceptible al
    // escribir pero absorbe la rafaga de teclas de una palabra completa.
    var searchDebounceTimer = null;
    sheetSearch.addEventListener('input', function(){
      if(sheetSearchClear){ sheetSearchClear.style.display = sheetSearch.value ? 'flex' : 'none'; }
      if(searchDebounceTimer){ clearTimeout(searchDebounceTimer); }
      searchDebounceTimer = setTimeout(runSearch, 130);
    });

    function runSearch(){
      var raw = sheetSearch.value.trim();
      var term = raw.toLowerCase();
      var tokens = term.split(/\s+/).filter(Boolean);
      if(!tokens.length){
        var existingEmpty = document.querySelector('#search-results .no-results');
        if(existingEmpty){ existingEmpty.remove(); }
        if(document.getElementById('search-results').classList.contains('active')){
          activate('home');
        }
        return;
      }
      renderResults(raw, tokens);
      activate('search-results');
    }

    if(sheetSearchClear){
      sheetSearchClear.addEventListener('click', function(){
        if(searchDebounceTimer){ clearTimeout(searchDebounceTimer); searchDebounceTimer = null; }
        sheetSearch.value = '';
        sheetSearchClear.style.display = 'none';
        var existingEmpty = document.querySelector('#search-results .no-results');
        if(existingEmpty){ existingEmpty.remove(); }
        if(document.getElementById('search-results').classList.contains('active')){
          activate('home', {keepSearchOpen:true});
        }
        if(searchBarRow){ searchBarRow.classList.add('open'); }
        sheetSearch.focus();
      });
    }

    // Initial load
    var initial = (location.hash || '').replace('#','') || 'home';
    if(!document.getElementById(initial)){ initial = 'home'; }
    activate(initial);

    // ===== Generador de Códigos de Barras EAN-13 / UPC-A en SVG estándar =====
    function parseEAN13(code) {
      var raw = String(code).replace(/\D/g, '');
      var s = raw;
      if (s.length === 12) { s = '0' + s; }
      if (s.length !== 13) return null;
      var sum = 0;
      for (var i = 0; i < 12; i++) {
        sum += parseInt(s[i], 10) * (i % 2 === 0 ? 1 : 3);
      }
      var check = (10 - (sum % 10)) % 10;
      if (check !== parseInt(s[12], 10)) { s = s.slice(0, 12) + check; }

      var L = ["0001101","0011001","0010011","0111101","0100011","0110001","0101111","0111011","0110111","0001011"];
      var G = ["0100111","0110011","0011011","0100001","0011101","0111001","0000101","0010001","0001001","0010111"];
      var R = ["1110010","1100110","1101100","1000010","1011100","1001110","1010000","1000100","1001000","1110100"];
      var P = ["LLLLLL","LLGLGG","LLGGLG","LLGGGL","LGLLGG","LGGLLG","LGGGLL","LGLGLG","LGLGGL","LGGLGL"];

      var first = parseInt(s[0], 10);
      var pattern = P[first];
      var leftBits = "";
      for (var j = 1; j <= 6; j++) {
        var dL = parseInt(s[j], 10);
        leftBits += (pattern[j - 1] === 'L') ? L[dL] : G[dL];
      }
      var rightBits = "";
      for (var k = 7; k <= 12; k++) {
        var dR = parseInt(s[k], 10);
        rightBits += R[dR];
      }
      return { firstDigit: s[0], leftDigits: s.slice(1, 7), rightDigits: s.slice(7, 13), leftBits: leftBits, rightBits: rightBits, isUPCA: raw.length === 12 };
    }

    function renderBarcodeSVG(rawCode) {
      var svg = document.getElementById('barcodeSvg');
      if (!svg) return;
      var clean = String(rawCode).trim();
      var ean = parseEAN13(clean);
      var moduleWidth = 2.4;
      var quietLeft = 11 * moduleWidth;
      var quietRight = 9 * moduleWidth;
      var barcodeModules = 95;
      var totalWidth = quietLeft + (barcodeModules * moduleWidth) + quietRight;
      var totalHeight = 115;
      var barHeightNormal = 75;
      var barHeightGuard = 87;

      svg.setAttribute('viewBox', '0 0 ' + totalWidth + ' ' + totalHeight);
      svg.setAttribute('width', '100%');
      svg.setAttribute('height', 'auto');
      svg.style.maxWidth = '340px';

      if (!ean) {
        svg.innerHTML = '<rect width="' + totalWidth + '" height="' + totalHeight + '" fill="#ffffff"/>' +
          '<text x="' + (totalWidth / 2) + '" y="' + (totalHeight / 2) + '" font-family="monospace" font-size="14" text-anchor="middle" fill="#111318">' + clean + '</text>';
        return;
      }

      var out = '<rect width="' + totalWidth + '" height="' + totalHeight + '" fill="#ffffff"/>';
      var currentX = quietLeft;

      function drawModules(bits, isGuard) {
        var h = isGuard ? barHeightGuard : barHeightNormal;
        for (var i = 0; i < bits.length; i++) {
          if (bits[i] === '1') {
            out += '<rect x="' + currentX.toFixed(2) + '" y="8" width="' + moduleWidth.toFixed(2) + '" height="' + h + '" fill="#000000"/>';
          }
          currentX += moduleWidth;
        }
      }

      drawModules("101", true);
      drawModules(ean.leftBits, false);
      drawModules("01010", true);
      drawModules(ean.rightBits, false);
      drawModules("101", true);

      var textY = 104;
      var fontSize = 14;
      var monoFont = "'SF Mono', 'Menlo', 'Consolas', monospace";
      out += '<text x="' + (quietLeft - (moduleWidth * 3.5)).toFixed(2) + '" y="' + textY + '" font-family="' + monoFont + '" font-size="' + fontSize + '" font-weight="700" text-anchor="middle" fill="#000000">' + ean.firstDigit + '</text>';
      var leftGroupCenter = quietLeft + (3 * moduleWidth) + (21 * moduleWidth);
      out += '<text x="' + leftGroupCenter.toFixed(2) + '" y="' + textY + '" font-family="' + monoFont + '" font-size="' + fontSize + '" font-weight="600" letter-spacing="2" text-anchor="middle" fill="#000000">' + ean.leftDigits + '</text>';
      var rightGroupCenter = quietLeft + ((3 + 42 + 5) * moduleWidth) + (21 * moduleWidth);
      out += '<text x="' + rightGroupCenter.toFixed(2) + '" y="' + textY + '" font-family="' + monoFont + '" font-size="' + fontSize + '" font-weight="600" letter-spacing="2" text-anchor="middle" fill="#000000">' + ean.rightDigits + '</text>';
      out += '<text x="' + (totalWidth - (moduleWidth * 3)).toFixed(2) + '" y="' + textY + '" font-family="' + monoFont + '" font-size="12" font-weight="700" text-anchor="middle" fill="#000000">&gt;</text>';
      svg.innerHTML = out;
    }

    var barcodeModal = document.getElementById('barcodeModal');
    var barcodeCloseBtn = document.getElementById('barcodeCloseBtn');
    var barcodeShareBtn = document.getElementById('barcodeShareBtn');
    var barcodeModalTitle = document.getElementById('barcodeModalTitle');
    var barcodeProductImgWrap = document.getElementById('barcodeProductImgWrap');
    var barcodeProductImg = document.getElementById('barcodeProductImg');

    // Muestra la foto del producto si existe en pictures/<hoja>/<ean>.webp.
    // La mayoria de referencias no tienen imagen, asi que ANTES de tocar la red
    // se consulta IMAGE_MANIFEST (generado en el build escaneando pictures/):
    // si el EAN no esta listado para esa hoja, se oculta el hueco al instante
    // sin disparar ninguna peticion 404. Solo existe .webp — no hay .jpeg.
    function hasProductImage(sheetFolder, eanCode) {
      if (!sheetFolder || !eanCode) return false;
      var bucket = IMAGE_MANIFEST[sheetFolder];
      return !!bucket && bucket.indexOf(eanCode) !== -1;
    }

    function productImageUrl(sheetFolder, eanCode) {
      return 'pictures/' + encodeURIComponent(sheetFolder) + '/' + encodeURIComponent(eanCode) + '.webp';
    }

    // Cache de objetos Image ya descargados/decodificados, por URL. El cuello de
    // botella percibido no era el peso del fichero (~13 KB) sino que la descarga
    // no empezaba hasta el clic: el usuario veia el barcode al instante (SVG
    // sincrono) y la foto llegaba un round-trip mas tarde. Precargando en cuanto
    // hay intencion (hover/touch) la imagen suele estar lista antes del clic.
    var IMG_CACHE = {};

    function prefetchProductImage(sheetFolder, eanCode) {
      if (!hasProductImage(sheetFolder, eanCode)) return null;
      var url = productImageUrl(sheetFolder, eanCode);
      if (IMG_CACHE[url]) return IMG_CACHE[url];
      var img = new Image();
      IMG_CACHE[url] = img;
      // decoding=async + decode() dejan el bitmap listo fuera del hilo critico,
      // para que al mostrarlo no haya un salto por decodificar en el momento.
      img.decoding = 'async';
      img.src = url;
      if (img.decode) { img.decode().catch(function(){}); }
      return img;
    }

    function updateBarcodeProductImage(sheetFolder, eanCode) {
      if (!barcodeProductImgWrap || !barcodeProductImg) return;
      if (!hasProductImage(sheetFolder, eanCode)) {
        barcodeProductImgWrap.style.display = 'none';
        barcodeProductImg.removeAttribute('src');
        return;
      }
      var url = productImageUrl(sheetFolder, eanCode);
      var cached = prefetchProductImage(sheetFolder, eanCode);

      // Camino rapido: si el prefetch ya termino, se muestra en este mismo frame
      // sin esperar ningun evento — el navegador sirve el bitmap desde cache.
      if (cached && cached.complete && cached.naturalWidth > 0) {
        barcodeProductImg.onload = null;
        barcodeProductImg.onerror = null;
        barcodeProductImg.src = url;
        barcodeProductImgWrap.style.display = 'flex';
        return;
      }

      // Camino lento: aun descargando. Se deja el hueco oculto hasta que cargue,
      // igual que antes, pero la descarga ya iba en marcha desde el hover.
      barcodeProductImgWrap.style.display = 'none';
      barcodeProductImg.onload = function() { barcodeProductImgWrap.style.display = 'flex'; };
      barcodeProductImg.onerror = function() {
        barcodeProductImgWrap.style.display = 'none';
        barcodeProductImg.removeAttribute('src');
      };
      barcodeProductImg.src = url;
    }

    function openBarcodeModal(eanCode, sheetFolder, eciValue) {
      if (!eanCode || !barcodeModal) return;
      if (barcodeModalTitle) {
        barcodeModalTitle.textContent = eciValue || '';
      }
      renderBarcodeSVG(eanCode);
      updateBarcodeProductImage(sheetFolder, eanCode);
      setModalOpen(barcodeModal, true);
    }

    // ===== Compartir la ventana de codigo de barras (referencia + foto +
    // codigo de barras) como una unica imagen PNG, via Web Share API con
    // fallback a descarga directa donde no hay share nativo (escritorio). =====
    function svgToCanvas(svgEl, targetWidth, cb) {
      if (!svgEl) { cb(null); return; }
      var vb = (svgEl.getAttribute('viewBox') || '0 0 300 150').split(' ');
      var w = parseFloat(vb[2]) || 300;
      var h = parseFloat(vb[3]) || 150;
      var scale = targetWidth / w;
      var clone = svgEl.cloneNode(true);
      clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
      clone.setAttribute('width', w);
      clone.setAttribute('height', h);
      var svgStr = new XMLSerializer().serializeToString(clone);
      var svgBlob = new Blob([svgStr], { type: 'image/svg+xml;charset=utf-8' });
      var url = URL.createObjectURL(svgBlob);
      var img = new Image();
      img.onload = function() {
        var canvas = document.createElement('canvas');
        canvas.width = w * scale;
        canvas.height = h * scale;
        var ctx = canvas.getContext('2d');
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        URL.revokeObjectURL(url);
        cb(canvas);
      };
      img.onerror = function() { URL.revokeObjectURL(url); cb(null); };
      img.src = url;
    }

    // Compone en un unico canvas, de arriba a abajo: referencia ECI, foto del
    // producto (si la hay) y el codigo de barras — con las mismas proporciones
    // relativas que se ven en la ventana modal (misma escala para todos los
    // elementos, solo renderizado a mayor resolucion para que no se vea
    // pixelado) — para que el PNG compartido sea autocontenido.
    function buildShareCanvas(cb) {
      var title = (barcodeModalTitle && barcodeModalTitle.textContent) || '';
      var hasImg = !!(barcodeProductImgWrap && barcodeProductImgWrap.style.display !== 'none' &&
        barcodeProductImg && barcodeProductImg.complete && barcodeProductImg.naturalWidth > 0);
      // Escala respecto al ancho real de la tarjeta del modal en pantalla, para
      // que el codigo de barras no salga mas grande de lo que se ve ahi.
      var cardWidth = (function() {
        var card = barcodeModal && barcodeModal.querySelector('.barcode-modal-card');
        return (card && card.clientWidth) || 340;
      })();
      var exportScale = 3;
      var padding = Math.round(22 * exportScale);
      var contentWidth = Math.round((cardWidth - 44) * exportScale);
      svgToCanvas(document.getElementById('barcodeSvg'), contentWidth, function(barcodeCanvas) {
        if (!barcodeCanvas) { cb(null); return; }
        var titleFontSize = Math.round(20 * exportScale);
        var titleHeight = title ? Math.round(titleFontSize * 1.5) : 0;
        var imgMaxHeight = hasImg ? Math.round(220 * exportScale) : 0;
        var gap = hasImg ? Math.round(14 * exportScale) : 0;
        var canvas = document.createElement('canvas');
        canvas.width = contentWidth + padding * 2;
        canvas.height = padding * 2 + titleHeight + imgMaxHeight + gap + barcodeCanvas.height;
        var ctx = canvas.getContext('2d');
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        var y = padding;
        if (title) {
          ctx.fillStyle = '#000000';
          ctx.font = '700 ' + titleFontSize + 'px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif';
          ctx.textAlign = 'center';
          ctx.textBaseline = 'top';
          ctx.fillText(title, canvas.width / 2, y);
          y += titleHeight;
        }
        function drawBarcodeAndFinish() {
          var bx = (canvas.width - barcodeCanvas.width) / 2;
          ctx.drawImage(barcodeCanvas, bx, y);
          cb(canvas);
        }
        if (hasImg) {
          var pImg = new Image();
          pImg.onload = function() {
            var maxW = contentWidth;
            var maxH = imgMaxHeight;
            var ratio = Math.min(maxW / pImg.naturalWidth, maxH / pImg.naturalHeight, 1);
            var dw = pImg.naturalWidth * ratio;
            var dh = pImg.naturalHeight * ratio;
            ctx.drawImage(pImg, (canvas.width - dw) / 2, y + (imgMaxHeight - dh) / 2, dw, dh);
            y += imgMaxHeight + gap;
            drawBarcodeAndFinish();
          };
          pImg.onerror = drawBarcodeAndFinish;
          pImg.src = barcodeProductImg.src;
        } else {
          drawBarcodeAndFinish();
        }
      });
    }

    function shareBarcodeModal() {
      if (!barcodeShareBtn || barcodeShareBtn.disabled) return;
      barcodeShareBtn.disabled = true;
      buildShareCanvas(function(canvas) {
        if (!canvas) { barcodeShareBtn.disabled = false; return; }
        canvas.toBlob(function(blob) {
          barcodeShareBtn.disabled = false;
          if (!blob) return;
          var eciText = (barcodeModalTitle && barcodeModalTitle.textContent) || 'referencia';
          var fileName = (eciText.replace(/[^a-z0-9]+/gi, '-').toLowerCase() || 'referencia') + '.png';
          var file;
          try { file = new File([blob], fileName, { type: 'image/png' }); } catch (err) { file = null; }
          if (file && navigator.canShare && navigator.canShare({ files: [file] })) {
            navigator.share({ files: [file] }).catch(function(){});
            return;
          }
          // Sin Web Share API con ficheros (tipicamente escritorio): se
          // descarga el PNG para que el usuario lo adjunte el mismo donde
          // quiera (mensaje, correo...).
          var url = URL.createObjectURL(blob);
          var a = document.createElement('a');
          a.href = url;
          a.download = fileName;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          setTimeout(function(){ URL.revokeObjectURL(url); }, 4000);
        }, 'image/png');
      });
    }
    if (barcodeShareBtn) barcodeShareBtn.addEventListener('click', shareBarcodeModal);

    function closeBarcodeModal() {
      setModalOpen(barcodeModal, false);
    }

    if (barcodeCloseBtn) barcodeCloseBtn.addEventListener('click', closeBarcodeModal);
    if (barcodeModal) {
      barcodeModal.addEventListener('click', function(e) {
        if (e.target === barcodeModal) closeBarcodeModal();
      });
    }
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape' && barcodeModal && barcodeModal.classList.contains('open')) {
        closeBarcodeModal();
      }
    });

    function setupEanClickables() {
      document.querySelectorAll('.panel:not(#search-results) table tbody tr').forEach(function(tr) {
        var cells = tr.querySelectorAll('td');
        if (cells.length >= 4) {
          var eanCell = cells[3];
          var eanSpan = eanCell.querySelector('.cell-text') || eanCell;
          var val = eanSpan.textContent.trim();
          if (val && /^\d{8,14}$/.test(val) && !tr.dataset.eanBound) {
            tr.dataset.eanBound = 'true';
            // Precarga la foto en cuanto el usuario roza la fila, para que al
            // hacer clic la imagen ya este en cache y aparezca junto al barcode
            // en vez de un round-trip mas tarde. Solo toca la red si el
            // manifiesto dice que esa referencia tiene foto.
            var prefetchRow = function() {
              var panelEl = tr.closest('.panel');
              prefetchProductImage(panelEl ? panelEl.dataset.sheetFolder : null, val);
            };
            tr.addEventListener('mouseenter', prefetchRow);
            tr.addEventListener('touchstart', prefetchRow, { passive: true });
            // Modelo, P/N, ECI y EAN abren el mismo modal: cualquiera de los
            // 4 primeros campos identifica igual de bien la referencia.
            [cells[0], cells[1], cells[2], cells[3]].forEach(function(cell) {
              if (!cell) return;
              var span = cell.querySelector('.cell-text') || cell;
              span.classList.add('ean-clickable');
              span.title = 'Ver código de barras';
              span.addEventListener('click', function(e) {
                e.stopPropagation();
                var panelEl = tr.closest('.panel');
                var sheetFolder = panelEl ? panelEl.dataset.sheetFolder : null;
                var eciCell = cells[2];
                var eciSpan = eciCell ? (eciCell.querySelector('.cell-text') || eciCell) : null;
                var eciValue = eciSpan ? eciSpan.textContent.trim() : '';
                openBarcodeModal(val, sheetFolder, eciValue);
              });
            });
          }
        }
      });

      var searchResults = document.getElementById('searchResultsBody');
      if (searchResults && !searchResults.dataset.eanBound) {
        searchResults.dataset.eanBound = 'true';
        // Mismo prefetch que en las tablas, por delegacion: las filas de
        // resultados se regeneran en cada busqueda, asi que no se puede enlazar
        // fila por fila.
        var prefetchFromEvent = function(e) {
          var rowTr = e.target.closest ? e.target.closest('tr') : null;
          if (!rowTr || !rowTr.dataset.ean) return;
          prefetchProductImage(rowTr.dataset.sheetFolder, rowTr.dataset.ean);
        };
        searchResults.addEventListener('mouseover', prefetchFromEvent);
        searchResults.addEventListener('touchstart', prefetchFromEvent, { passive: true });
        searchResults.addEventListener('click', function(e) {
          var target = e.target.closest('.result-cell-inner');
          if (target) {
            var valueSpan = target.querySelector('.result-cell-value');
            if (valueSpan && valueSpan.classList.contains('ean-clickable')) {
              var rowTr = target.closest('tr');
              var eanVal = rowTr ? rowTr.dataset.ean : '';
              if (eanVal && /^\d{8,14}$/.test(eanVal)) {
                e.stopPropagation();
                var sheetFolder = rowTr ? rowTr.dataset.sheetFolder : null;
                var eciValue = rowTr ? rowTr.dataset.eci : '';
                openBarcodeModal(eanVal, sheetFolder, eciValue);
              }
            }
          }
        });
      }
    }

    setupEanClickables();

    // ===== Identificar producto por foto (OCR) =====
    var scanBtn = document.getElementById('scanToggle');
    var scannerModal = document.getElementById('scannerModal');
    var scannerCloseBtn = document.getElementById('scannerCloseBtn');
    var scannerStatus = document.getElementById('scannerStatus');
    var scannerResults = document.getElementById('scannerResults');
    var scannerDetected = document.getElementById('scannerDetected');
    var scannerDetectedText = document.getElementById('scannerDetectedText');
    var scannerLiveWrap = document.getElementById('scannerLiveWrap');
    var scannerVideo = document.getElementById('scannerVideo');
    var scannerLiveOverlay = document.getElementById('scannerLiveOverlay');
    var scannerLiveStopBtn = document.getElementById('scannerLiveStopBtn');
    var scannerLiveFrame = document.getElementById('scannerLiveFrame');
    var scannerLiveCrop = document.getElementById('scannerLiveCrop');
    var scannerBarcodeOverlay = document.getElementById('scannerBarcodeOverlay');
    var scannerZoomRow = document.getElementById('scannerZoomRow');
    var scannerZoomSlider = document.getElementById('scannerZoomSlider');
    var scannerZoomValue = document.getElementById('scannerZoomValue');

    // El escaner en vivo esta pensado para la camara trasera de un movil, no
    // para la webcam de un portatil: mostrarlo en escritorio solo confunde y
    // exhibe el video flotando sobre la pagina sin ningun caso de uso real.
    // isMobileDevice() combina el sniffing clasico de UA con un heuristico
    // moderno para iPadOS 13+, que se anuncia como escritorio pero es tactil.
    function isMobileDevice() {
      var ua = navigator.userAgent || '';
      if (/Android|iPhone|iPad|iPod|Mobile/i.test(ua)) return true;
      // iPadOS 13+ en modo escritorio: touch points > 1 y plataforma Mac.
      if (/Macintosh/i.test(ua) && navigator.maxTouchPoints > 1) return true;
      return false;
    }

    if (scanBtn && isMobileDevice() && hasGetUserMedia()) { scanBtn.style.display = 'flex'; }
    // Barcode detector init is intentionally NOT called here: on Safari/iOS
    // it triggers a network fetch for the WASM ponyfill, and running that on
    // every page load would tax first-visit load time for no benefit until
    // the user actually scans. It is kicked off lazily by the same
    // hover/touchstart prewarm as the OCR worker, and again (idempotently)
    // by startBarcodeFastPath() when the camera actually opens.

    // --- Live scanner state ---
    var liveStream = null;
    var liveTrack = null;
    var liveTimer = null;
    var liveBusy = false;
    var liveActive = false;
    var liveZoom = 1.0;              // digital zoom factor applied to the crop
    var liveNativeZoom = null;       // {min,max,step} when the camera exposes optical/native zoom
    var LIVE_INTERVAL_MS = 700;

    // --- Captura a resolucion de sensor ------------------------------------
    // El recorte que se le pasa al OCR solo puede tener los pixeles que traiga
    // el frame de origen, asi que la resolucion del track es el techo duro de
    // legibilidad: con un preview de 1280 una referencia ECI de estanteria cae
    // a ~10px de alto y Tesseract no puede leerla por mucho que se reescale.
    // Se pide el maximo que acepte la camara y, si el track declara un maximo
    // mayor del concedido, se sube en caliente con applyConstraints.
    var liveImageCapture = null;     // ImageCapture: existe en Chrome/Android, no en Safari/iOS
    var liveCaptureCanvas = null;    // canvas reutilizado para el detect del barcode (downscale)

    // --- Localizacion de la etiqueta (ROI) ----------------------------------
    // Un codigo de barras es un patron de rayas verticales de frecuencia muy
    // alta, y eso se detecta contando cruces de gradiente por fila SIN
    // decodificar nada. Importa porque BarcodeDetector solo devuelve
    // boundingBox cuando ha conseguido decodificar, y el caso que interesa
    // aqui es justo el contrario: etiqueta gastada o borrosa que no decodifica,
    // donde hay que leer por OCR los digitos impresos. Teniendo la banda de
    // barras localizada, el legible esta pegado debajo (a veces encima), asi
    // que el recorte se ancla ahi en vez de al centro ciego del encuadre.
    var liveRoi = null;              // {x,y,w,h} normalizado 0..1 sobre el frame
    var liveRoiAt = 0;               // timestamp de la ultima deteccion
    var liveRoiMiss = 0;             // ticks seguidos sin lectura con ROI fijada
    var ROI_TTL_MS = 1600;           // pasado este tiempo se vuelve al recuadro central
    var ROI_TILES_X = 20;
    var ROI_TILES_Y = 14;
    var ROI_WORK_W = 360;            // ancho de trabajo del detector de rayas
    var roiWorkCanvas = null;
    // Alternates the OCR alphabet between ticks (digits-only / part-number) and
    // keeps a short history of reads so both alphabets contribute to one match.
    var liveProfileToggle = false;
    var liveTextWindow = [];

    // A single persistent Tesseract worker. Tesseract.recognize() spins up and
    // tears down a worker on EVERY call (1-3s each), which made most live ticks
    // no-ops. Reusing one worker also makes the digit whitelist actually apply —
    // the convenience API silently ignores tessedit_* options.
    var ocrWorker = null;
    var ocrWorkerPromise = null;
    var ocrWorkerMode = null;

    function getOcrWorker() {
      if (ocrWorkerPromise) return ocrWorkerPromise;
      if (typeof Tesseract === 'undefined' || !Tesseract.createWorker) return null;
      ocrWorkerPromise = Tesseract.createWorker()
        .then(function(w) {
          return w.loadLanguage('eng').then(function() {
            return w.initialize('eng');
          }).then(function() {
            ocrWorker = w;
            return w;
          });
        }).catch(function(err) {
          ocrWorkerPromise = null;
          throw err;
        });
      return ocrWorkerPromise;
    }

    // Two recognition profiles. The single mixed alphabet used before was the
    // reason ECI codes failed while P/N worked: allowing A-Z lets Tesseract pick
    // a letter for every ambiguous digit (5->S, 0->O, 1->I, 8->B), and an ECI is
    // pure digits, so one letter anywhere destroys the whole 11-digit run. A P/N
    // legitimately contains letters, hence it was unaffected.
    var OCR_PROFILES = {
      // Digits only — for ECI and EAN. Tesseract cannot emit a letter at all.
      // Separators stay allowed: the extractor uses them to rebuild grouped
      // codes like "0 195950 638813", and banning them makes Tesseract invent
      // digits where a gap really is.
      digits: '0123456789. -',
      // Digits + the P/N alphabet (B, I, O and S excluded: they never occur in a
      // real part number, so banning them removes the confusions at the source).
      pn: '0123456789/ ACDEFGHJKLMNPQRTUVWXYZ'
    };

    function setOcrProfile(worker, profile, pageseg) {
      var key = profile + '|' + pageseg;
      if (ocrWorkerMode === key) return Promise.resolve();
      ocrWorkerMode = key;
      return worker.setParameters({
        tessedit_char_whitelist: OCR_PROFILES[profile],
        tessedit_pageseg_mode: pageseg
      }).catch(function() {});
    }

    function hasGetUserMedia() {
      return !!(navigator.mediaDevices && navigator.mediaDevices.getUserMedia);
    }

    function openScanner() {
      if (!scannerModal) return;
      setModalOpen(scannerModal, true);
      resetScannerUI();
      if (hasGetUserMedia()) {
        startLiveScan();
      } else {
        scannerStatus.textContent = 'Este navegador no permite acceso a la camara. Usa el buscador de la pagina.';
        scannerStatus.className = 'scanner-status error';
      }
    }

    function closeScanner() {
      if (!scannerModal) return;
      setModalOpen(scannerModal, false);
      stopLiveScan();
    }

    function resetScannerUI() {
      scannerResults.innerHTML = '';
      scannerStatus.textContent = '';
      scannerStatus.className = 'scanner-status';
      scannerDetected.style.display = 'none';
    }

    // Restart scanning after a result: clear the previous hit and reopen the camera.
    function restartLiveScan() {
      resetScannerUI();
      stopLiveScan();
      startLiveScan();
    }

    // --- Live scanning: capture a frame every LIVE_INTERVAL_MS, run OCR on it,
    // never overlap two OCR passes, and stop automatically on an exact match.
    function startLiveScan() {
      if (liveActive) return;
      if (!hasGetUserMedia()) {
        scannerLiveOverlay.textContent = 'Camara no disponible en este navegador.';
        return;
      }
      liveActive = true;
      liveTextWindow = [];
      scannerLiveWrap.style.display = 'block';
      scannerLiveStopBtn.style.display = 'block';
      scannerZoomRow.classList.add('visible');
      scannerLiveOverlay.textContent = 'Iniciando camara...';

      // Se pide la resolucion mas alta que la camara acepte: el recorte es lo
      // que se OCR'ea, asi que los pixeles de origen determinan directamente
      // cuan legibles son los digitos pequenos. 3840 de ancho es lo que dan
      // hoy tanto los iPhone (Safari respeta width.ideal) como los Android de
      // gama media; si la camara no llega, la propia negociacion de
      // getUserMedia devuelve el modo mas cercano por debajo, no un error.
      var constraints = {
        video: {
          facingMode: { ideal: 'environment' },
          width: { ideal: 3840 },
          height: { ideal: 2160 },
          focusMode: { ideal: 'continuous' }
        }
      };

      // Start the OCR worker in parallel with the camera — it is the slow part.
      // Que no este disponible NO es motivo para abortar el escaneo: el lector
      // de codigos de barras no depende de Tesseract para nada, asi que se
      // sigue adelante en modo "solo barcode" y lo unico que se pierde es la
      // lectura de los digitos impresos.
      var workerPromise = getOcrWorker();
      if (!workerPromise) {
        scannerLiveOverlay.textContent = 'OCR no disponible: solo lectura de codigo de barras.';
      }

      navigator.mediaDevices.getUserMedia(constraints)
        .catch(function() {
          // Escalon intermedio antes de rendirse: muchas webcam de portatil y
          // algun Android antiguo rechazan 3840 pero si dan 1920, y perder esa
          // opcion significaba caer al modo por defecto (640x480 en algunos
          // navegadores), que es ilegible para OCR.
          return navigator.mediaDevices.getUserMedia({
            video: { facingMode: { ideal: 'environment' }, width: { ideal: 1920 }, height: { ideal: 1080 } }
          });
        })
        .catch(function() {
          // Fall back to a plain request if the ideal constraints are rejected.
          return navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
        })
        .then(function(stream) {
          if (!liveActive) { stream.getTracks().forEach(function(t){ t.stop(); }); return null; }
          liveStream = stream;
          liveTrack = stream.getVideoTracks()[0] || null;
          detectNativeZoom();
          detectCaptureCapabilities();
          scannerVideo.srcObject = stream;
          return scannerVideo.play();
        }).then(function() {
          if (!liveActive) return;
          // La subida de resolucion va DESPUES de play() y a proposito: una
          // renegociacion del track tarda unas decimas y, puesta antes, deja la
          // vista en negro ese tiempo. El elemento <video> absorbe el cambio de
          // tamano en caliente, asi que el usuario ve imagen de inmediato y los
          // pixeles extra entran por debajo sin interrumpir nada.
          maximiseTrackResolution();
          // La camara se revela EN CUANTO hay imagen. Antes se esperaba ademas
          // a que Tesseract estuviera cargado (~10 MB de modelo, 1-3 s) para no
          // mostrar nunca una vista "a medias". El efecto real en tienda era el
          // contrario del buscado: el usuario miraba una pantalla de espera
          // mientras el lector de codigos de barras -- que no necesita OCR para
          // nada -- ya estaba corriendo debajo y podia haber resuelto el
          // producto. Con un codigo legible el escaneo termina muchas veces
          // antes de que el OCR acabe de cargar, asi que bloquear la vista solo
          // anadia segundos por producto.
          startBarcodeFastPath();
          scannerLiveWrap.classList.add('camera-ready');
          applyZoom(liveZoom);
          scannerLiveOverlay.textContent = workerPromise
            ? 'Apunta al codigo de barras (preparando lectura de digitos)...'
            : 'Apunta al codigo de barras...';
          if (!workerPromise) return;

          // El OCR se incorpora cuando esta listo, sin bloquear la camara. Si
          // falla, el escaneo sigue vivo leyendo codigos de barras en vez de
          // caerse entero.
          workerPromise.then(function() {
            if (!liveActive) return;
            scannerLiveOverlay.textContent = 'Encuadra el numero dentro del recuadro...';
            liveTimer = setInterval(liveScanTick, LIVE_INTERVAL_MS);
          }).catch(function() {
            if (!liveActive) return;
            scannerLiveOverlay.textContent = 'Solo codigo de barras: no se pudo cargar el OCR.';
          });
        }).catch(function(err) {
          scannerLiveOverlay.textContent = 'No se pudo iniciar la camara. Revisa los permisos.';
          stopLiveScan();
        });
    }

    // --- Native barcode fast path -------------------------------------------
    // The EAN on a label is printed as an actual barcode symbol, not just as
    // digits. Where the browser exposes the native BarcodeDetector API (the
    // same class of API DemoUnit's live rectangle-highlight relies on), decode
    // the symbol directly instead of OCR'ing the human-readable digits below
    // it: it runs every animation frame (not throttled to LIVE_INTERVAL_MS),
    // it does not need Tesseract to be ready, and it cannot misread a glyph
    // because it never converts pixels to characters. Safari/iOS has no native
    // support today, so this silently degrades to the existing OCR pipeline —
    // it never replaces it, only wins the race when it can.
    var barcodeDetector = null;
    var barcodeSupported = false;
    var barcodeIsPolyfill = false;   // true when running the WASM ponyfill (Safari/iOS)
    var barcodeRafId = null;
    var barcodeTimerId = null;
    var POLYFILL_TICK_MS = 180;      // throttle for the WASM path — a bare rAF loop is too
                                      // heavy for older iPhones running OCR at the same time.
    var barcodeReadyPromise = null;  // memoized: never load the polyfill script twice.

    // Loads a BarcodeDetector-compatible ponyfill ONLY when the native API is
    // absent (Safari/iOS today) and ONLY when actually triggered — either by
    // hovering/touching the scan button (prewarm) or by opening the scanner
    // directly. Chrome/Android never pays this cost: window.BarcodeDetector
    // already exists there and this resolves synchronously.
    function ensureBarcodeDetector() {
      if (barcodeReadyPromise) return barcodeReadyPromise;
      if (typeof window.BarcodeDetector !== 'undefined') {
        barcodeIsPolyfill = false;
        barcodeReadyPromise = Promise.resolve(true);
        return barcodeReadyPromise;
      }
      barcodeReadyPromise = new Promise(function(resolve) {
        var settled = false;
        function finish(ok) { if (settled) return; settled = true; resolve(ok); }

        function onReady() {
          window.removeEventListener('barcodedetector-polyfill-ready', onReady);
          barcodeIsPolyfill = true;
          finish(true);
        }
        window.addEventListener('barcodedetector-polyfill-ready', onReady);

        // module script so we can `import` the ESM build; assigns onto window
        // so the rest of the code never needs to know native vs. polyfill.
        var script = document.createElement('script');
        script.type = 'module';
        script.textContent =
          "import { BarcodeDetectorPolyfill } from 'https://esm.sh/barcode-detector@2/pure';" +
          "window.BarcodeDetector = BarcodeDetectorPolyfill;" +
          "window.dispatchEvent(new Event('barcodedetector-polyfill-ready'));";
        script.onerror = function() { finish(false); };
        document.head.appendChild(script);

        // Safety net: a slow/broken CDN load must never block the camera or
        // the OCR pipeline — after 4s, give up on the fast path for this
        // session and let OCR carry the whole scan as it always could.
        setTimeout(function() { finish(typeof window.BarcodeDetector !== 'undefined'); }, 4000);
      });
      return barcodeReadyPromise;
    }

    function initBarcodeDetector() {
      return ensureBarcodeDetector().then(function(ok) {
        if (!ok) { barcodeSupported = false; return false; }
        try {
          // ean_13/upc_a cubren el codigo de la caja del producto; code_128 y
          // code_39 son los que llevan las etiquetas de estanteria de ECI, que
          // codifican la referencia ECI de 11 digitos. En tienda esa etiqueta
          // se escanea tanto o mas que la propia caja.
          // ean_8 y upc_e se han quitado a proposito: codifican 8 digitos y en
          // el catalogo no hay ninguna referencia de esa longitud, asi que solo
          // anadian trabajo de decodificacion por frame sin poder acertar nunca.
          barcodeDetector = new window.BarcodeDetector({ formats: ['ean_13', 'upc_a', 'code_128', 'code_39'] });
          barcodeSupported = true;
          return true;
        } catch (e) { barcodeSupported = false; return false; }
      });
    }

    function drawBarcodeBoxes(boxes) {
      if (!scannerBarcodeOverlay) return;
      var vw = scannerVideo.videoWidth, vh = scannerVideo.videoHeight;
      if (!vw || !vh) return;
      if (scannerBarcodeOverlay.width !== vw) scannerBarcodeOverlay.width = vw;
      if (scannerBarcodeOverlay.height !== vh) scannerBarcodeOverlay.height = vh;
      var ctx = scannerBarcodeOverlay.getContext('2d');
      ctx.clearRect(0, 0, vw, vh);
      ctx.lineWidth = Math.max(2, vw * 0.004);
      ctx.strokeStyle = 'rgba(52,199,89,.95)';
      boxes.forEach(function(b) {
        if (!b) return;
        ctx.strokeRect(b.x, b.y, b.width, b.height);
      });
    }

    // Runs on requestAnimationFrame for the native API (cheap enough to do
    // every frame). The WASM ponyfill is throttled to POLYFILL_TICK_MS instead
    // — decoding in WASM is real CPU work, and running it unthrottled would
    // compete with the OCR loop on older iPhones.
    var BARCODE_DETECT_W = 1600;   // ancho maximo al que se le entrega el frame al detector

    // Pedir 4K al track mejora el OCR pero encarece el decode: a 3840x2160 y en
    // cada animation frame el detector -- y mucho mas el ponyfill WASM -- se come
    // el hilo y arrastra consigo al bucle de OCR. Las barras se decodifican de
    // sobra a 1600px de ancho, asi que el decode va sobre una copia reducida y
    // los pixeles de verdad se reservan para el recorte que lee Tesseract, que
    // es donde si hacen falta. Las cajas se devuelven a coordenadas de video
    // multiplicando por el factor de reduccion.
    function barcodeDetectSource() {
      var vw = scannerVideo.videoWidth, vh = scannerVideo.videoHeight;
      if (!vw || vw <= BARCODE_DETECT_W) return { src: scannerVideo, k: 1 };
      var dw = BARCODE_DETECT_W;
      var dh = Math.max(1, Math.round(vh * dw / vw));
      if (!liveCaptureCanvas) liveCaptureCanvas = document.createElement('canvas');
      if (liveCaptureCanvas.width !== dw || liveCaptureCanvas.height !== dh) {
        liveCaptureCanvas.width = dw;
        liveCaptureCanvas.height = dh;
      }
      try {
        liveCaptureCanvas.getContext('2d').drawImage(scannerVideo, 0, 0, dw, dh);
      } catch (e) { return { src: scannerVideo, k: 1 }; }
      return { src: liveCaptureCanvas, k: vw / dw };
    }

    function barcodeFastTick() {
      if (!liveActive || !barcodeSupported || !scannerVideo.videoWidth) return;
      var frame = barcodeDetectSource();
      barcodeDetector.detect(frame.src).then(function(codes) {
        if (!liveActive) return;
        var k = frame.k;
        var boxes = codes.map(function(c) {
          var b = c.boundingBox;
          if (!b || k === 1) return b;
          return { x: b.x * k, y: b.y * k, width: b.width * k, height: b.height * k };
        });
        drawBarcodeBoxes(boxes);

        // Un barcode decodificado entrega la ROI exacta y gratis (ya esta
        // calculada). Donde mas importa es cuando el valor leido NO esta en el
        // catalogo: tipicamente la etiqueta lleva el EAN del fabricante en
        // barras y la referencia util (ECI, P/N) impresa al lado en claro. Antes
        // ese caso era un callejon sin salida; ahora la caja ancla el recorte y
        // el OCR ataca la banda adyacente en vez del centro del encuadre.
        if (!roiIsFresh() && boxes.length && scannerVideo.videoWidth) {
          var bb = boxes[0];
          if (bb && bb.width > 0 && bb.height > 0) {
            setLiveRoi({
              x: bb.x / scannerVideo.videoWidth,
              y: bb.y / scannerVideo.videoHeight,
              w: bb.width / scannerVideo.videoWidth,
              h: bb.height / scannerVideo.videoHeight
            });
          }
        }

        var unknown = null;
        for (var i = 0; i < codes.length; i++) {
          var value = (codes[i].rawValue || '').replace(/\D/g, '');
          // Misma clasificacion por longitud que classifyRef: 11 digitos es una
          // referencia ECI (la etiqueta de estanteria, en code_128), 12 y 13
          // son EAN. El 13 pierde el cero inicial porque asi se guarda en el
          // indice del catalogo.
          var cl = null;
          if (value.length === 11) cl = { kind: 'ECI', value: value };
          else if (value.length === 12) cl = { kind: 'EAN', value: value };
          else if (value.length === 13) cl = { kind: 'EAN', value: value.replace(/^0/, '') };
          if (!cl) continue;
          var hit = matchExactRef(cl.kind, cl.value);
          if (hit) {
            stopLiveScan(true);
            scannerStatus.textContent = '1 producto encontrado';
            scannerStatus.className = 'scanner-status found';
            scannerDetectedText.textContent = cl.kind + ': ' + cl.value;
            scannerDetected.style.display = 'block';
            displayScanResults([hit], false);
            return;
          }
          unknown = cl;
        }
        // Un codigo decodificado correctamente pero que no esta dado de alta
        // producia silencio absoluto: el bucle seguia girando y parecia que el
        // lector no funcionaba, cuando en realidad habia leido perfectamente.
        // Decirlo evita que el usuario insista contra un producto que no esta
        // en el catalogo.
        if (unknown) {
          var msg = unknown.kind + ' ' + unknown.value + ' leido — no esta en el catalogo';
          if (scannerStatus.textContent !== msg) {
            scannerStatus.textContent = msg;
            scannerStatus.className = 'scanner-status';
          }
        }
        scheduleBarcodeTick();
      }).catch(function() {
        scheduleBarcodeTick();
      });
    }

    function scheduleBarcodeTick() {
      if (!liveActive || !barcodeSupported) return;
      if (barcodeIsPolyfill) {
        barcodeTimerId = setTimeout(barcodeFastTick, POLYFILL_TICK_MS);
      } else {
        barcodeRafId = requestAnimationFrame(barcodeFastTick);
      }
    }

    function startBarcodeFastPath() {
      // initBarcodeDetector() is async (it may still be loading the polyfill
      // script). Guard with liveActive so a scan that was stopped/restarted
      // before the load finished does not start a stray loop.
      initBarcodeDetector().then(function(ok) {
        if (!ok || !liveActive) return;
        drawBarcodeBoxes([]);
        scheduleBarcodeTick();
      });
    }

    function stopBarcodeFastPath() {
      if (barcodeRafId) { cancelAnimationFrame(barcodeRafId); barcodeRafId = null; }
      if (barcodeTimerId) { clearTimeout(barcodeTimerId); barcodeTimerId = null; }
      drawBarcodeBoxes([]);
    }

    // Exact-match lookup against the same REF_INDEX the OCR path scores
    // against, but with no fuzzy edit-distance pass — a decoded barcode is
    // either right or it failed the checksum, so there is nothing to score.
    function matchExactRef(kind, value) {
      var bucket = getRefIndex()[kind];
      if (!bucket) return null;
      for (var i = 0; i < bucket.length; i++) {
        if (bucket[i].cmp === value) {
          return { row: bucket[i].row, kind: kind, ref: bucket[i].display, token: bucket[i].display, read: value, dist: 0, sim: 1 };
        }
      }
      return null;
    }

    // Some cameras expose real zoom through MediaStreamTrack capabilities. iOS
    // Safari usually does not, hence the digital-crop fallback in applyZoom.
    function detectNativeZoom() {
      liveNativeZoom = null;
      if (!liveTrack || !liveTrack.getCapabilities) return;
      try {
        var caps = liveTrack.getCapabilities();
        if (caps && caps.zoom && typeof caps.zoom.max === 'number' && caps.zoom.max > caps.zoom.min) {
          liveNativeZoom = caps.zoom;
        }
      } catch (e) { liveNativeZoom = null; }
    }

    // ImageCapture solo existe en Chrome/Android. Donde esta, grabFrame()
    // entrega el fotograma del track directamente como ImageBitmap, sin pasar
    // por el elemento <video> ni por su escalado CSS. En Safari/iOS no existe
    // y no hay sustituto: alli se dibuja el <video>, que a resolucion de track
    // alta da practicamente lo mismo. takePhoto() se ha descartado a
    // proposito: da el still del sensor completo, pero con latencia de
    // obturador y sonido de disparo en Android, inaceptable en un bucle de
    // tienda donde se escanean referencias en cadena.
    function detectCaptureCapabilities() {
      liveImageCapture = null;
      if (!liveTrack) return;
      try {
        if (typeof window.ImageCapture === 'function') {
          liveImageCapture = new window.ImageCapture(liveTrack);
        }
      } catch (e) { liveImageCapture = null; }
    }

    // getUserMedia negocia UN modo y se queda ahi, pero el track suele declarar
    // un maximo mayor del concedido (tipico: concede 1920 y puede 4032). Subirlo
    // en caliente es gratis en calidad de lectura y es el unico camino en
    // Safari/iOS, donde no hay ImageCapture para pedir el still del sensor.
    function maximiseTrackResolution() {
      if (!liveTrack || !liveTrack.getCapabilities || !liveTrack.applyConstraints) return;
      var caps;
      try { caps = liveTrack.getCapabilities(); } catch (e) { return; }
      if (!caps || !caps.width || typeof caps.width.max !== 'number') return;
      var current = 0;
      try { current = (liveTrack.getSettings() || {}).width || 0; } catch (e) {}
      // Margen del 20%: renegociar para ganar unos pocos pixeles cuesta mas
      // (parpadeo, reenfoque) de lo que aporta.
      if (caps.width.max <= current * 1.2) return;
      try {
        liveTrack.applyConstraints({ width: { ideal: caps.width.max } })
          .then(function() { detectNativeZoom(); })
          .catch(function() {});
      } catch (e) {}
    }

    // Devuelve una fuente dibujable con el frame actual al maximo detalle
    // disponible, como promesa porque grabFrame() es asincrono. El <video> es
    // el fallback universal.
    function grabStillSource() {
      if (liveImageCapture && liveImageCapture.grabFrame) {
        return liveImageCapture.grabFrame().catch(function() { return scannerVideo; });
      }
      return Promise.resolve(scannerVideo);
    }

    // Zoom has two independent effects and both help legibility:
    //  - native zoom (when available) makes the sensor itself zoom in;
    //  - the crop region shrinks, so the same OCR budget covers fewer, bigger pixels.
    function applyZoom(z) {
      liveZoom = z;
      scannerZoomValue.textContent = z.toFixed(1) + 'x';

      if (liveNativeZoom && liveTrack) {
        var span = liveNativeZoom.max - liveNativeZoom.min;
        // Map the first half of the slider range onto the native zoom range;
        // beyond that the digital crop takes over.
        var frac = Math.min(1, (z - 1) / 3);
        var target = liveNativeZoom.min + span * frac;
        try { liveTrack.applyConstraints({ advanced: [{ zoom: target }] }); } catch (e) {}
      }

      // Shrink the guide frame in step with the crop so what the user frames is
      // exactly what gets analysed. Con una etiqueta ya localizada el recuadro
      // lo posiciona setLiveRoi() y no se toca aqui: el zoom del slider solo
      // manda mientras el recorte es el central ciego.
      if (roiIsFresh()) return;
      var wPct = 90 / z;
      var hPct = 55 / z;
      scannerLiveFrame.style.width = wPct + '%';
      scannerLiveFrame.style.height = hPct + '%';
      scannerLiveFrame.style.left = ((100 - wPct) / 2) + '%';
      scannerLiveFrame.style.top = ((100 - hPct) / 2) + '%';
    }

    // keepRestart: after a successful read the camera stops, but the user must
    // still be able to scan the next label without closing the modal.
    function stopLiveScan(keepRestart) {
      liveActive = false;
      liveBusy = false;
      stopBarcodeFastPath();
      if (liveTimer) { clearInterval(liveTimer); liveTimer = null; }
      if (liveStream) {
        liveStream.getTracks().forEach(function(t) { t.stop(); });
        liveStream = null;
      }
      liveTrack = null;
      liveImageCapture = null;
      liveRoi = null;
      liveRoiAt = 0;
      liveRoiMiss = 0;
      if (scannerLiveFrame) scannerLiveFrame.classList.remove('locked');
      if (scannerVideo) scannerVideo.srcObject = null;
      if (scannerLiveWrap) { scannerLiveWrap.style.display = 'none'; scannerLiveWrap.classList.remove('camera-ready'); }
      if (scannerLiveStopBtn) scannerLiveStopBtn.style.display = keepRestart ? 'block' : 'none';
      if (scannerZoomRow) scannerZoomRow.classList.remove('visible');
      if (scannerLiveCrop) scannerLiveCrop.textContent = '';
    }

    // --- Detector de banda de codigo de barras (sin decodificar) ------------
    // Mide, sobre una copia reducida en gris, cuantas veces cambia el signo del
    // gradiente horizontal en cada fila. Una banda de barras da una tasa de
    // cruces alta Y casi identica en todas sus filas; un bloque de texto da una
    // tasa parecida en las filas de glifos pero se derrumba en las de
    // interlineado, asi que la UNIFORMIDAD VERTICAL es lo que de verdad separa
    // barras de texto, no la tasa en si. Es aritmetica de canvas: cuesta una
    // fraccion de un tick de Tesseract y no depende de ninguna API.
    function detectBarcodeBand(src, sw, sh) {
      if (!sw || !sh) return null;
      var W = Math.min(ROI_WORK_W, sw);
      var H = Math.max(8, Math.round(sh * W / sw));
      if (!roiWorkCanvas) roiWorkCanvas = document.createElement('canvas');
      if (roiWorkCanvas.width !== W || roiWorkCanvas.height !== H) {
        roiWorkCanvas.width = W;
        roiWorkCanvas.height = H;
      }
      var ctx = roiWorkCanvas.getContext('2d');
      var data;
      try {
        ctx.drawImage(src, 0, 0, W, H);
        data = ctx.getImageData(0, 0, W, H).data;
      } catch (e) { return null; }

      var gray = new Uint8Array(W * H);
      for (var p = 0, g = 0; p < data.length; p += 4, g++) {
        gray[g] = (data[p] * 0.299 + data[p + 1] * 0.587 + data[p + 2] * 0.114) | 0;
      }

      var tw = Math.floor(W / ROI_TILES_X);
      var th = Math.floor(H / ROI_TILES_Y);
      if (tw < 4 || th < 3) return null;
      var DELTA = 18;        // cambio minimo de gris para contar como transicion
      var RATE_MIN = 0.16;   // cruces por pixel: una transicion cada ~6px
      var CV_MAX = 0.40;     // desviacion/media maxima entre filas del tile
      var grid = new Uint8Array(ROI_TILES_X * ROI_TILES_Y);

      for (var ty = 0; ty < ROI_TILES_Y; ty++) {
        for (var tx = 0; tx < ROI_TILES_X; tx++) {
          var x0 = tx * tw, y0 = ty * th;
          var sum = 0, sumSq = 0, rows = 0;
          for (var y = y0; y < y0 + th; y++) {
            var base = y * W;
            var crossings = 0, lastSign = 0;
            for (var x = x0; x < x0 + tw - 1; x++) {
              var d = gray[base + x + 1] - gray[base + x];
              if (d > DELTA) {
                if (lastSign === -1) crossings++;
                lastSign = 1;
              } else if (d < -DELTA) {
                if (lastSign === 1) crossings++;
                lastSign = -1;
              }
            }
            var rate = crossings / tw;
            sum += rate;
            sumSq += rate * rate;
            rows++;
          }
          if (!rows) continue;
          var mean = sum / rows;
          if (mean < RATE_MIN) continue;
          var variance = Math.max(0, sumSq / rows - mean * mean);
          if (Math.sqrt(variance) / mean > CV_MAX) continue;
          grid[ty * ROI_TILES_X + tx] = 1;
        }
      }

      var best = largestTileCluster(grid, ROI_TILES_X, ROI_TILES_Y);
      if (!best || best.count < 3) return null;
      return {
        x: (best.x0 * tw) / W,
        y: (best.y0 * th) / H,
        w: ((best.x1 - best.x0 + 1) * tw) / W,
        h: ((best.y1 - best.y0 + 1) * th) / H
      };
    }

    // Componente conexa (4-vecinos) mas grande del grid de tiles marcados, con
    // su bounding box. Se coge la mayor y no todas las marcadas juntas porque en
    // una estanteria entran varias etiquetas en el encuadre y unir sus cajas
    // daria un recorte que abarca toda la escena, justo lo que se quiere evitar.
    function largestTileCluster(grid, cols, rows) {
      var seen = new Uint8Array(cols * rows);
      var best = null;
      var stack = [];
      for (var i = 0; i < grid.length; i++) {
        if (!grid[i] || seen[i]) continue;
        stack.length = 0;
        stack.push(i);
        seen[i] = 1;
        var count = 0, x0 = cols, x1 = -1, y0 = rows, y1 = -1;
        while (stack.length) {
          var cur = stack.pop();
          var cx = cur % cols, cy = (cur - cx) / cols;
          count++;
          if (cx < x0) x0 = cx;
          if (cx > x1) x1 = cx;
          if (cy < y0) y0 = cy;
          if (cy > y1) y1 = cy;
          var nb = [
            cx > 0 ? cur - 1 : -1,
            cx < cols - 1 ? cur + 1 : -1,
            cy > 0 ? cur - cols : -1,
            cy < rows - 1 ? cur + cols : -1
          ];
          for (var n = 0; n < 4; n++) {
            var idx = nb[n];
            if (idx >= 0 && grid[idx] && !seen[idx]) { seen[idx] = 1; stack.push(idx); }
          }
        }
        if (!best || count > best.count) best = { count: count, x0: x0, x1: x1, y0: y0, y1: y1 };
      }
      return best;
    }

    // Expande la banda de barras hasta abarcar el legible impreso. Se incluyen
    // las dos bandas adyacentes a proposito: en la caja del producto el EAN en
    // claro va DEBAJO de las barras, pero en la etiqueta de estanteria de ECI la
    // referencia suele ir ENCIMA, y a priori no se sabe cual de las dos se esta
    // apuntando.
    function roiToCropRect(roi) {
      var padX = roi.w * 0.12;
      var x = roi.x - padX;
      var w = roi.w + padX * 2;
      var y = roi.y - roi.h * 0.50;
      var h = roi.h * 2.20;
      // Suelo de tamano: un recorte demasiado estrecho deja a Tesseract sin
      // contexto de linea y empeora la lectura en vez de mejorarla.
      if (w < 0.35) { x -= (0.35 - w) / 2; w = 0.35; }
      if (h < 0.14) { y -= (0.14 - h) / 2; h = 0.14; }
      if (x < 0) x = 0;
      if (y < 0) y = 0;
      if (x + w > 1) w = 1 - x;
      if (y + h > 1) h = 1 - y;
      return { x: x, y: y, w: w, h: h };
    }

    function setLiveRoi(roi) {
      liveRoi = roi;
      liveRoiAt = roi ? Date.now() : 0;
      liveRoiMiss = 0;
      // Cambiar de etiqueta invalida la ventana de lecturas: mezclar digitos de
      // dos etiquetas distintas es la forma mas facil de construir una
      // referencia que no existe y darla por buena.
      liveTextWindow = [];
      if (!scannerLiveFrame) return;
      if (roi) {
        var r = roiToCropRect(roi);
        scannerLiveFrame.style.left = (r.x * 100) + '%';
        scannerLiveFrame.style.top = (r.y * 100) + '%';
        scannerLiveFrame.style.width = (r.w * 100) + '%';
        scannerLiveFrame.style.height = (r.h * 100) + '%';
        scannerLiveFrame.classList.add('locked');
      } else {
        scannerLiveFrame.classList.remove('locked');
        applyZoom(liveZoom);
      }
    }

    function roiIsFresh() {
      return !!(liveRoi && (Date.now() - liveRoiAt) < ROI_TTL_MS);
    }


    // Recorta la region a analizar y la reescala HACIA ARRIBA a un ancho de
    // trabajo fijo. Si hay una banda de barras localizada y reciente el recorte
    // se ancla a ella -- que es lo que hace que los digitos llenen el recorte en
    // vez de ocupar una franja perdida dentro de una foto de estanteria; si no,
    // se vuelve al recuadro guia centrado, el comportamiento que ya habia.
    function grabLiveCrop(src, sw, sh) {
      if (!sw || !sh) return null;

      var anchored = roiIsFresh();
      var cropX, cropY, cropW, cropH;
      if (anchored) {
        var r = roiToCropRect(liveRoi);
        cropX = Math.round(r.x * sw);
        cropY = Math.round(r.y * sh);
        cropW = Math.round(r.w * sw);
        cropH = Math.round(r.h * sh);
      } else {
        cropW = Math.round(sw * 0.90 / liveZoom);
        cropH = Math.round(sh * 0.55 / liveZoom);
        cropX = Math.round((sw - cropW) / 2);
        cropY = Math.round((sh - cropH) / 2);
      }
      if (cropW < 16 || cropH < 8) return null;

      // Con el recorte anclado la region es bastante mas pequena, asi que el
      // mismo presupuesto de pixeles se gasta entero en los digitos: se puede
      // subir el ancho de trabajo y el tope de ampliacion sin alargar el tick.
      var TARGET_W = anchored ? 1600 : 1400;
      var MAX_SCALE = anchored ? 5 : 3;
      var scale = Math.min(MAX_SCALE, Math.max(1, TARGET_W / cropW));
      var outW = Math.round(cropW * scale);
      var outH = Math.round(cropH * scale);

      var canvas = document.createElement('canvas');
      canvas.width = outW;
      canvas.height = outH;
      var ctx = canvas.getContext('2d');
      ctx.imageSmoothingEnabled = true;
      if ('imageSmoothingQuality' in ctx) ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(src, cropX, cropY, cropW, cropH, 0, 0, outW, outH);
      preprocessCrop(ctx, outW, outH);

      scannerLiveCrop.textContent = outW + 'x' + outH + (anchored ? ' - fijado' : '');
      return canvas;
    }

    // Grayscale + contrast stretch. Label print is high-contrast dark-on-light,
    // so this mostly removes the noise that makes Tesseract hallucinate digits.
    function preprocessCrop(ctx, outW, outH) {
      try {
        var img = ctx.getImageData(0, 0, outW, outH);
        var d = img.data;
        var i;
        var lo = 255, hi = 0;
        for (i = 0; i < d.length; i += 4) {
          var g = (d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114) | 0;
          d[i] = d[i + 1] = d[i + 2] = g;
          if (g < lo) lo = g;
          if (g > hi) hi = g;
        }
        var range = hi - lo;
        if (range > 20) {
          var mul = 255 / range;
          for (i = 0; i < d.length; i += 4) {
            var v = (d[i] - lo) * mul;
            v = v < 0 ? 0 : (v > 255 ? 255 : v);
            d[i] = d[i + 1] = d[i + 2] = v;
          }
        }
        ctx.putImageData(img, 0, 0);
      } catch (e) { /* tainted canvas should not happen with a local stream */ }
    }

    function liveScanTick() {
      if (liveBusy || !liveActive) return;
      if (!ocrWorker) return;
      if (!scannerVideo.videoWidth) return;
      liveBusy = true;

      grabStillSource().then(function(src) {
        if (!liveActive) { liveBusy = false; return; }
        var sw = src.videoWidth || src.width || 0;
        var sh = src.videoHeight || src.height || 0;

        // Se relocaliza la etiqueta solo cuando no hay ROI vigente: mientras el
        // usuario mantiene el encuadre, repetir la deteccion cada tick haria
        // bailar el recorte entre fotogramas y las lecturas de la ventana
        // (liveTextWindow) dejarian de corresponder a la misma etiqueta.
        if (!roiIsFresh()) {
          var band = detectBarcodeBand(src, sw, sh);
          if (band) setLiveRoi(band);
        }

        var canvas = grabLiveCrop(src, sw, sh);
        if (src && src.close) src.close();   // ImageBitmap: liberar el buffer
        if (!canvas) { liveBusy = false; return; }
        return runOcrOnCrop(canvas);
      }).catch(function() {
        liveBusy = false;
      });
    }

    function runOcrOnCrop(canvas) {
      // PSM 6 = "single uniform block of text", the right model for a label crop.
      // Two passes over the SAME crop with different alphabets. A single mixed
      // alphabet is a bad compromise: it lets letters win on digit-only codes.
      // Alternating profiles also spreads the cost — each tick does one pass.
      liveProfileToggle = !liveProfileToggle;
      var profile = liveProfileToggle ? 'digits' : 'pn';

      return setOcrProfile(ocrWorker, profile, '6').then(function() {
        return ocrWorker.recognize(canvas);
      }).then(function(result) {
        liveBusy = false;
        if (!liveActive) return;
        var text = (result && result.data && result.data.text) || '';

        // Keep the last few reads from both profiles, so a digit-pass ECI and a
        // pn-pass P/N found on consecutive ticks are evaluated together.
        liveTextWindow.push(text);
        if (liveTextWindow.length > 3) liveTextWindow.shift();

        var outcome = evaluateOcrText(liveTextWindow.join('\n'));
        if (outcome.ambiguous) {
          var dedupedAmb = dedupeMatchesByRow(outcome.matches);
          scannerLiveOverlay.textContent = 'Leido: ' + outcome.reads.slice(0, 2).join(' | ') + ' — varias coincidencias posibles';
          scannerStatus.textContent = dedupedAmb.length + ' posibles coincidencias — confirma visualmente cual es';
          scannerStatus.className = 'scanner-status';
          scannerDetectedText.textContent = outcome.reads.slice(0, 5).join(' | ');
          scannerDetected.style.display = 'block';
          displayScanResults(outcome.matches, true);
          return;
        }
        if (outcome.matches.length === 0) {
          // Una ROI puede ser un falso positivo: un bloque de texto denso
          // (ingredientes, legal) tambien da tasa de cruces alta y uniforme. Si
          // estando fijado no sale ninguna lectura en dos ticks seguidos, se
          // suelta el anclaje para que el detector vuelva a buscar en vez de
          // insistir sobre una zona equivocada hasta que caduque el TTL.
          if (roiIsFresh() && !outcome.reads.length) {
            liveRoiMiss++;
            if (liveRoiMiss >= 2) setLiveRoi(null);
          } else {
            liveRoiMiss = 0;
          }
          scannerLiveOverlay.textContent = outcome.reads.length
            ? 'Leido: ' + outcome.reads.slice(0, 2).join(' | ') + ' — sin coincidencia, acerca o ajusta el zoom...'
            : (roiIsFresh()
                ? 'Etiqueta localizada, leyendo los digitos...'
                : 'Buscando referencia... usa el zoom si los digitos se ven pequenos');
          return;
        }
        if (outcome.exact) {
          stopLiveScan(true);
          var dedupedExact = dedupeMatchesByRow(outcome.matches);
          scannerStatus.textContent = dedupedExact.length + ' producto' + (dedupedExact.length > 1 ? 's' : '') + ' encontrado' + (dedupedExact.length > 1 ? 's' : '');
          scannerStatus.className = 'scanner-status found';
          scannerDetectedText.textContent = outcome.reads.slice(0, 5).join(' | ');
          scannerDetected.style.display = 'block';
          displayScanResults(outcome.matches, false);
        } else {
          scannerLiveOverlay.textContent = 'Posible coincidencia (1 errata) — sigue encuadrando para confirmar';
          scannerStatus.textContent = 'Sin lectura exacta. Hay una unica referencia a una errata de lo leido; comprueba el codigo antes de darla por buena.';
          scannerStatus.className = 'scanner-status';
          scannerDetectedText.textContent = outcome.reads.slice(0, 5).join(' | ');
          scannerDetected.style.display = 'block';
          displayScanResults(outcome.matches, true);
        }
      }).catch(function() {
        liveBusy = false;
      });
    }

    // OCR confuses letters for digits inside numeric runs. Aggressive mapping,
    // used only when re-reading the text as digits.
    function ocrDigitFix(s) {
      return s
        .replace(/[OoQD]/g, '0')
        .replace(/[IilL|!\]\[]/g, '1')
        .replace(/[Zz]/g, '2')
        .replace(/[Ss]/g, '5')
        .replace(/[Gg]/g, '6')
        .replace(/[Tt]/g, '7')
        .replace(/[Bb]/g, '8')
        .replace(/[gq]/g, '9');
    }

    // Apple part numbers never contain B, I, O or S (verified against the whole
    // catalogue), so those glyphs in a P/N read are certainly misread digits.
    function pnFix(s) {
      return s.replace(/O/g, '0').replace(/I/g, '1').replace(/S/g, '5').replace(/B/g, '8');
    }

    // Reduce a raw value to the canonical form both sides are compared in.
    // Returns null when the value is not a recognisable reference.
    function classifyRef(raw) {
      var s = (raw || '').trim();
      if (!s) return null;
      if (/^[\d.\s]+$/.test(s)) {
        var d = s.replace(/\D/g, '');
        if (d.length === 11) return { kind: 'ECI', value: d };
        if (d.length === 12 || d.length === 13) return { kind: 'EAN', value: d.replace(/^0/, '') };
        return null;
      }
      var up = s.toUpperCase();
      if (/^[A-Z0-9]{4,9}\/[A-Z]{1,2}$/.test(up)) {
        return { kind: 'PN', value: up.replace(/[^A-Z0-9]/g, '') };
      }
      return null;
    }

    // Digito de control del EAN-13. El catalogo guarda los EAN sin el cero
    // inicial (classifyRef hace replace(/^0/,'')), y un UPC-A de 12 digitos es
    // un EAN-13 con cero delante y el mismo digito de control, asi que basta
    // con rellenar por la izquierda hasta 13 y aplicar la formula estandar.
    function eanChecksumOk(value) {
      var d = String(value || '').replace(/\D/g, '');
      if (!d || d.length > 13) return false;
      while (d.length < 13) d = '0' + d;
      var sum = 0;
      for (var i = 0; i < 12; i++) sum += (+d.charAt(i)) * (i % 2 === 0 ? 1 : 3);
      return ((10 - (sum % 10)) % 10) === (+d.charAt(12));
    }

    // Pure extraction + matching, shared by photo mode and live mode. Never
    // touches the DOM, so live mode can call it every tick without side effects.
    // Patron de separador entre grupos de digitos de un ECI impreso ("152.154
    // 27335", "152-154-27335"...). Es fijo, asi que las dos regex que lo usan se
    // compilan UNA vez a nivel de modulo en vez de en cada tick del escaner
    // (~700ms, en paralelo con Tesseract): recrearlas ahi era trabajo identico
    // repetido sin ningun dato nuevo del que depender.
    var OCR_ECI_SEP = '[.,:;\\u00b7\\-\\s]';
    // ECI 3/3/4/1 — the variant seen on some labels. Must be tried BEFORE
    // the 3/3/5 rule, whose prefix it shares.
    var RE_ECI_3341 = new RegExp('(\\d{3})\\s*' + OCR_ECI_SEP + '\\s*(\\d{3})\\s*' + OCR_ECI_SEP + '\\s*(\\d{4})\\s*' + OCR_ECI_SEP + '\\s*(\\d)(?!\\d)', 'g');
    // ECI 3/3/5 — the common form.
    var RE_ECI_335 = new RegExp('(\\d{3})\\s*' + OCR_ECI_SEP + '\\s*(\\d{3})\\s*' + OCR_ECI_SEP + '\\s*(\\d{5})(?!\\d)', 'g');

    function evaluateOcrText(rawText) {
      var empty = { tokens: [], matches: [], exact: false, reads: [] };
      if (!rawText || !rawText.trim()) return empty;

      var tokens = [];
      var seen = {};

      function addToken(value, kind) {
        if (!value) return;
        // OCR sometimes reads the black/white bars of a barcode symbol itself
        // (not the human-readable digits/characters printed below it) as a
        // long run of a single repeated character -- almost always "1"s from
        // the bars' vertical strokes. No real EAN/ECI/PN in the catalogue
        // repeats one character 5+ times in a row, so treat that shape as a
        // barcode-symbol misread and discard it outright instead of letting
        // it chase fuzzy matches.
        if (/(.)\1{4,}/.test(value)) return;
        var key = kind + ':' + value;
        if (value.length >= 5 && !seen[key]) {
          seen[key] = true;
          tokens.push({ value: value, kind: kind });
        }
      }

      var up = rawText.toUpperCase();

      // --- Numeric references: read the text twice, once as-is and once with
      // letter->digit corrections, so a single bad glyph cannot hide a reference.
      [up, ocrDigitFix(up)].forEach(function(text) {
        // Las regex son 'g' y compartidas entre llamadas: hay que resetear
        // lastIndex o un exec() previo que no llegara al final del texto
        // dejaria la siguiente busqueda empezando a mitad de cadena.
        var m;
        RE_ECI_3341.lastIndex = 0;
        while ((m = RE_ECI_3341.exec(text)) !== null) addToken(m[1] + m[2] + m[3] + m[4], 'ECI');

        RE_ECI_335.lastIndex = 0;
        while ((m = RE_ECI_335.exec(text)) !== null) addToken(m[1] + m[2] + m[3], 'ECI');

        // Sliding windows over digit runs. Crucially the run is built by joining
        // digit groups across in-line separators: EAN under a barcode is printed
        // as "0 195950 638813", so requiring 12 CONTIGUOUS digits found nothing.
        // Newlines are excluded so unrelated lines are never welded together.
        var joinRe = /\d[\d .,\-·]*\d/g;
        var jm;
        while ((jm = joinRe.exec(text)) !== null) {
          var run = jm[0].replace(/\D/g, '');
          var i;
          for (i = 0; i + 11 <= run.length; i++) addToken(run.substr(i, 11), 'ECI');
          for (i = 0; i + 12 <= run.length; i++) addToken(run.substr(i, 12), 'EAN');
          for (i = 0; i + 13 <= run.length; i++) addToken(run.substr(i, 13).replace(/^0/, ''), 'EAN');
        }
      });

      // --- Part numbers: with the slash, and as a bare alphanumeric run in case
      // the slash was not picked up.
      var pnRe = /([A-Z0-9]{4,9})\s*\/\s*([A-Z]{1,2})/g;
      var pm;
      while ((pm = pnRe.exec(up)) !== null) addToken(pnFix(pm[1] + pm[2]), 'PN');
      (up.match(/[A-Z0-9]{6,10}/g) || []).forEach(function(run) {
        if (/[A-Z]/.test(run) && /\d/.test(run)) addToken(pnFix(run), 'PN');
      });

      if (tokens.length === 0) return empty;

      var refIndex = getRefIndex();

      // Lo que se muestra cuando no hay ninguna coincidencia fiable: las
      // lecturas mas largas, para que el usuario vea QUE se ha leido y pueda
      // juzgar si merece la pena repetir el escaneo.
      function readsFallback() {
        return tokens.slice()
          .sort(function(a, b) { return b.value.length - a.value.length; })
          .slice(0, 4)
          .map(function(t) { return t.kind + ': ' + t.value; });
      }

      // Tolerancia de lectura POR TIPO de referencia, en numero de erratas, en
      // lugar de un unico porcentaje de parecido para todo. Un porcentaje unico
      // enganaba: al depender de la longitud, el 80% se traducia en tolerar DOS
      // erratas en EAN y ECI y UNA en los PN, que es justo al reves de lo que
      // admite el catalogo. Medido sobre las referencias reales:
      //   - a UNA errata no hay ni una sola pareja de EAN (0 de 1101) ni de ECI
      //     (0 de 1705) que se confunda entre si, asi que esa errata se puede
      //     tolerar sin riesgo de cruzar dos productos, y es la que rescata una
      //     lectura buena con un solo glifo mal leido;
      //   - a DOS erratas el 82% de los EAN y el 99% de los ECI ya tienen otra
      //     referencia valida a esa distancia: de ahi venian los resultados
      //     equivocados, porque el escaner siempre encontraba "algo" plausible;
      //   - los PN son correlativos por color/capacidad (3P128ZFA / 3P129ZFA),
      //     asi que el 98% colisiona con una sola errata. Para ellos la unica
      //     tolerancia segura es cero.
      var MAX_DIST = { EAN: 1, ECI: 1, PN: 0 };

      var best = {};
      tokens.forEach(function(t) {
        var bucket = refIndex[t.kind];
        if (!bucket) return;
        var maxDist = MAX_DIST[t.kind];
        if (maxDist === undefined) return;
        // Un EAN cuyo digito de control cuadra es una lectura coherente consigo
        // misma: salvo casualidad muy improbable, es exactamente el codigo
        // impreso. Si no esta en el catalogo es que no esta, y corregirlo a un
        // vecino seria inventarselo (tipico al escanear un producto que no
        // tenemos dado de alta). Por eso solo se le admite coincidencia exacta.
        // Si el digito de control NO cuadra, la lectura si viene corrupta y ahi
        // corregir una errata tiene todo el sentido.
        if (t.kind === 'EAN' && eanChecksumOk(t.value)) maxDist = 0;
        bucket.forEach(function(entry) {
          var ref = entry.cmp;
          var dist;
          if (maxDist === 0) {
            if (ref !== t.value) return;
            dist = 0;
          } else {
            if (Math.abs(ref.length - t.value.length) > maxDist) return;
            dist = boundedEditDistance(ref, t.value, maxDist);
            if (dist < 0) return;
          }
          var sim = 1 - dist / Math.max(ref.length, t.value.length);
          var prev = best[entry.key];
          if (!prev || sim > prev.sim || (sim === prev.sim && dist < prev.dist)) {
            best[entry.key] = {
              row: entry.row, kind: t.kind, ref: entry.display,
              token: entry.display, read: t.value, dist: dist, sim: sim
            };
          }
        });
      });

      var results = Object.keys(best).map(function(k) { return best[k]; });

      if (results.length === 0) {
        return { tokens: tokens, matches: [], exact: false, reads: readsFallback() };
      }

      // Exact reads first, then by similarity.
      results.sort(function(a, b) { return b.sim - a.sim || a.ref.length - b.ref.length; });

      var exact = results.filter(function(r) { return r.dist === 0; });
      var shown;

      if (exact.length > 0) {
        shown = exact;
      } else {
        // Sin lectura exacta solo queda la correccion de una errata, y una
        // correccion solo es fiable si es UNICA. Si dos productos DISTINTOS
        // quedan a la misma distancia de lo leido no hay forma de saber cual
        // es, y ensenar la lista equivale a acertar por sorteo: es preferible
        // no dar resultado y que el usuario repita la lectura. (Varias entradas
        // de la MISMA fila no son ambiguedad: es el mismo producto coincidiendo
        // por mas de una columna, de ahi el dedupe por fila antes de contar.)
        var bestDist = results.reduce(function(acc, r) { return Math.min(acc, r.dist); }, Infinity);
        var atBest = dedupeMatchesByRow(results.filter(function(r) { return r.dist === bestDist; }));
        if (atBest.length !== 1) {
          // Mas de un producto DISTINTO empata a la misma distancia: no hay
          // forma fiable de elegir uno solo. En vez de descartarlo como si no
          // se hubiera leido nada, se devuelven los candidatos empatados con
          // ambiguous:true para que la UI se los ensene al usuario y sea el
          // quien confirme visualmente cual es (en vez de acertar por sorteo).
          var usedReadsAmb = [];
          var usedSeenAmb = {};
          atBest.forEach(function(r) {
            var k = r.kind + ':' + r.read;
            if (!usedSeenAmb[k]) { usedSeenAmb[k] = true; usedReadsAmb.push(r.kind + ': ' + r.read); }
          });
          return { tokens: tokens, matches: atBest, exact: false, ambiguous: true, reads: usedReadsAmb };
        }
        shown = atBest;
      }

      var usedReads = [];
      var usedSeen = {};
      shown.forEach(function(r) {
        var k = r.kind + ':' + r.read;
        if (!usedSeen[k]) { usedSeen[k] = true; usedReads.push(r.kind + ': ' + r.read); }
      });

      return { tokens: tokens, matches: shown, exact: exact.length > 0, ambiguous: false, reads: usedReads };
    }

    // Banded edit distance. Returns -1 as soon as it is provably above max,
    // so scanning the whole catalogue stays cheap.
    function boundedEditDistance(a, b, max) {
      if (a === b) return 0;
      if (Math.abs(a.length - b.length) > max) return -1;
      var prev = new Array(b.length + 1);
      var cur = new Array(b.length + 1);
      for (var j = 0; j <= b.length; j++) prev[j] = j;
      for (var i = 1; i <= a.length; i++) {
        cur[0] = i;
        var rowMin = cur[0];
        var lo = Math.max(1, i - max);
        var hi = Math.min(b.length, i + max);
        for (var k = 1; k < lo; k++) cur[k] = max + 1;
        for (var jj = lo; jj <= hi; jj++) {
          var cost = a.charAt(i - 1) === b.charAt(jj - 1) ? 0 : 1;
          cur[jj] = Math.min(
            prev[jj] + 1,
            cur[jj - 1] + 1,
            prev[jj - 1] + cost
          );
          if (cur[jj] < rowMin) rowMin = cur[jj];
        }
        for (var k2 = hi + 1; k2 <= b.length; k2++) cur[k2] = max + 1;
        if (rowMin > max) return -1;
        var tmp = prev; prev = cur; cur = tmp;
      }
      var res = prev[b.length];
      return res > max ? -1 : res;
    }

    // Index of every EAN / ECI / P-N value in the catalogue, in the same canonical
    // form the OCR reads are reduced to. Classified by the SHAPE of the value, not
    // by the column header — 111 references live in columns whose header is blank
    // or product-specific and were previously invisible to the scanner.
    var REF_INDEX = null;
    function getRefIndex() {
      if (REF_INDEX) return REF_INDEX;
      REF_INDEX = { EAN: [], ECI: [], PN: [] };
      var dedupe = {};
      ALL_ROWS.forEach(function(r) {
        var rowKey = r.slug + '|' + r.section;
        r.cells.forEach(function(c) {
          var cl = classifyRef(c.v);
          if (!cl) return;
          var key = rowKey + '|' + cl.kind + '|' + cl.value;
          if (dedupe[key]) return;
          dedupe[key] = true;
          REF_INDEX[cl.kind].push({
            // cmp: the string the reads are compared against. For part numbers the
            // ambiguous glyphs are folded on BOTH sides so the comparison is fair.
            cmp: cl.kind === 'PN' ? pnFix(cl.value) : cl.value,
            display: cl.value,
            row: r,
            key: key
          });
        });
      });
      return REF_INDEX;
    }

    // A single product row often matches on more than one column (its EAN AND
    // its ECI both hit, or the same value appears twice in the row) — that is
    // still one product, so callers collapse to one entry per row before
    // counting or rendering. `matches` is already sorted best-first (see
    // evaluateOcrText / matchExactRef callers), so the first entry kept per
    // row is also the best one.
    function dedupeMatchesByRow(matches) {
      var seenRow = {};
      var deduped = [];
      matches.forEach(function(m) {
        var rowKey = m.row.slug + '|' + m.row.section;
        if (seenRow[rowKey]) return;
        seenRow[rowKey] = true;
        deduped.push(m);
      });
      return deduped;
    }

    function displayScanResults(matches, isApprox) {
      scannerResults.innerHTML = '';
      var deduped = dedupeMatchesByRow(matches);
      var shown = deduped.slice(0, 10);
      shown.forEach(function(m) {
        var r = m.row;
        var item = document.createElement('div');
        item.className = 'scanner-result-item';
        var titleDiv = document.createElement('div');
        titleDiv.className = 'scanner-result-title';
        titleDiv.textContent = r.sheet + ' — ' + r.section;
        item.appendChild(titleDiv);

        if (isApprox && m.sim !== undefined && m.dist) {
          var note = document.createElement('div');
          note.className = 'scanner-approx-note';
          note.textContent = Math.round(m.sim * 100) + '% · ' + m.kind + ' ' + m.ref +
            ' (leido ' + m.read + ')';
          item.appendChild(note);
        }

        var detailDiv = document.createElement('div');
        detailDiv.className = 'scanner-result-detail';
        // Highlight the cell that produced the hit, comparing in the same canonical
        // form. Falls back to a substring test for free-text manual searches.
        var tok = m.token || '';
        var tokCanon = null;
        if (m.kind) {
          tokCanon = m.kind === 'PN' ? pnFix(tok.toUpperCase().replace(/[^A-Z0-9]/g, '')) : tok.replace(/\D/g, '');
        }
        r.cells.forEach(function(c) {
          if (!c.v) return;
          var chip = document.createElement('span');
          chip.className = 'scanner-detail-chip';
          if (tokCanon) {
            var cl = classifyRef(c.v);
            if (cl && cl.kind === m.kind) {
              var cellCanon = cl.kind === 'PN' ? pnFix(cl.value) : cl.value;
              if (cellCanon === tokCanon) chip.classList.add('scanner-detail-match');
            }
          } else if (tok && c.v.toLowerCase().indexOf(tok.toLowerCase()) !== -1) {
            chip.classList.add('scanner-detail-match');
          }
          chip.textContent = (c.label ? c.label + ': ' : '') + c.v;
          detailDiv.appendChild(chip);
        });
        item.appendChild(detailDiv);
        item.addEventListener('click', function() {
          closeScanner();
          if (sheetSearch) {
            // Search using the value exactly as it appears in the catalogue
            // (with dots / slash), not the canonical form used for comparison.
            var q = m.token;
            if (tokCanon) {
              for (var ci = 0; ci < r.cells.length; ci++) {
                var cc = r.cells[ci];
                var ccl = classifyRef(cc.v);
                if (!ccl || ccl.kind !== m.kind) continue;
                var ccanon = ccl.kind === 'PN' ? pnFix(ccl.value) : ccl.value;
                if (ccanon === tokCanon) { q = cc.v; break; }
              }
            }
            sheetSearch.value = q;
            sheetSearch.dispatchEvent(new Event('input'));
          }
        });
        scannerResults.appendChild(item);
      });
      if (deduped.length > 10) {
        var more = document.createElement('div');
        more.className = 'scanner-no-match';
        more.textContent = '... y ' + (deduped.length - 10) + ' mas. Toca un resultado para ver en el catalogo.';
        scannerResults.appendChild(more);
      }
    }

    // --- Event listeners ---
    if (scanBtn) {
      scanBtn.addEventListener('click', openScanner);
      // Precalienta el worker de OCR en cuanto el usuario muestra intencion de
      // escanear (hover en desktop, touchstart en movil), asi el arranque de
      // 1-3s de Tesseract ya esta consumido cuando openScanner() lo necesite.
      scanBtn.addEventListener('mouseenter', getOcrWorker);
      scanBtn.addEventListener('touchstart', getOcrWorker, { passive: true });
      // Mismo precalentado para el detector de barcode: en Chrome/Android
      // ensureBarcodeDetector() resuelve al instante (API nativa), en
      // Safari/iOS dispara ya la descarga del polyfill WASM.
      scanBtn.addEventListener('mouseenter', ensureBarcodeDetector);
      scanBtn.addEventListener('touchstart', ensureBarcodeDetector, { passive: true });
    }
    if (scannerCloseBtn) scannerCloseBtn.addEventListener('click', closeScanner);
    if (scannerLiveStopBtn) scannerLiveStopBtn.addEventListener('click', restartLiveScan);
    if (scannerZoomSlider) {
      scannerZoomSlider.addEventListener('input', function() {
        applyZoom(parseInt(scannerZoomSlider.value, 10) / 10);
      });
    }
    if (scannerModal) {
      scannerModal.addEventListener('click', function(e) {
        if (e.target === scannerModal) closeScanner();
      });
      document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && scannerModal.classList.contains('open')) closeScanner();
      });
    }
  }

  // --- Parse events.txt ---
  function parseEvents(raw) {
    var events = {};
    raw.split('\n').forEach(function(line) {
      line = line.trim();
      if (!line || line.charAt(0) === '#') return;
      var parts = line.split(' | ');
      if (parts.length < 3) return;
      var date = parts[0].trim();
      var year = parseInt(parts[1].trim(), 10);
      var text = parts.slice(2).join(' | ').trim();
      if (!events[date]) events[date] = [];
      events[date].push({ y: year, t: text });
    });
    return events;
  }

  // --- Load external resources then init ---
  Promise.all([
    fetch('data.json').then(function(r){ return r.json(); }),
    fetch('events.txt').then(function(r){ return r.text(); }),
    // El manifiesto es una optimizacion, no un requisito: si falla, se cae a un
    // objeto vacio y hasProductImage() devuelve false para todo (ningun hueco
    // de imagen, cero peticiones) en vez de romper el arranque.
    fetch('images.json').then(function(r){ return r.json(); }).catch(function(){ return {}; })
  ]).then(function(results){
    ALL_ROWS = results[0];
    APPLE_ON_THIS_DAY = parseEvents(results[1]);
    IMAGE_MANIFEST = results[2] || {};
    initApp();
  }).catch(function(err){
    console.error('Error cargando datos externos:', err);
    initApp();
  });
})();
</script>
</body>
</html>
"""

if __name__ == '__main__':
    main()