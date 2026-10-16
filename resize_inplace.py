import os
import re
from PIL import Image

ROOT = "pictures"
MAX_W = 700
TARGET_BYTES = 30 * 1024  # ~30 KB objetivo
TOLERANCE = 4 * 1024      # margen aceptable (+/- 4KB) antes de reducir más el tamaño
MIN_W = 120               # no reducir el ancho por debajo de esto
SOURCE_EXTS = (".jpg", ".jpeg", ".png", ".webp")

# Nombres que macOS / iCloud / los navegadores generan al bajar dos veces el
# mismo fichero: "<base> 2.webp", "<base> 3.webp", "<base> copia.webp",
# "<base> copia 2.webp", "<base> copy.webp". Siempre llevan un espacio antes
# del sufijo; un nombre de imagen bueno (solo el EAN) nunca lo lleva, asi que
# el espacio es la senal inequivoca de duplicado.
DUP_STEM_RE = re.compile(
    r'^(?P<base>.+?)[ ]+(?:\d+|cop(?:ia|y)(?:[ ]+\d+)?)$',
    re.IGNORECASE,
)


def consolidate_duplicates(root):
    """Elimina los duplicados "<base> N.webp" dejando una sola copia por imagen.

    Se ejecuta DESPUES de la conversion, cuando todo es ya .webp:
      - Si el fichero base ("<base>.webp") existe, el duplicado se BORRA.
      - Si no existe (el duplicado es la unica copia de esa imagen), se RESCATA
        renombrandolo a "<base>.webp" en vez de perderlo.
    Devuelve (borrados, rescatados) para el resumen.
    """
    deleted, rescued = [], []
    for dirpath, _dirnames, filenames in os.walk(root):
        webps = [f for f in filenames if f.lower().endswith('.webp')]
        stems = set(f[:-5] for f in webps)
        for fname in sorted(webps):
            stem = fname[:-5]
            m = DUP_STEM_RE.match(stem)
            if not m:
                continue
            base = m.group('base').strip()
            src = os.path.join(dirpath, fname)
            if base in stems:
                os.remove(src)
                stems.discard(stem)
                deleted.append(src)
            else:
                dst = os.path.join(dirpath, base + '.webp')
                os.rename(src, dst)
                stems.discard(stem)
                stems.add(base)
                rescued.append((src, dst))
    return deleted, rescued



def size_at_quality(img, width, quality, tmp_path):
    work = img if img.width == width else img.resize(
        (width, max(1, int(img.height * (width / img.width)))), Image.LANCZOS
    )
    work.save(tmp_path, "WEBP", quality=quality, method=6)
    return os.path.getsize(tmp_path)


def encode_to_target(img, target_bytes, tmp_path):
    """Busca la calidad (y si hace falta, ancho) que da el tamaño MAS CERCANO
    al objetivo, tanto por encima como por debajo. Solo reduce el ancho si
    incluso con calidad minima el archivo sigue pesando mas que el objetivo
    (nunca para "inflar" una imagen que ya es pequeña de por si)."""
    width = img.width
    best = None  # (diff, size, quality, width)

    while True:
        # Muestreo de calidades para encontrar el minimo local mas cercano al objetivo
        candidates = [1, 10, 20, 30, 40, 50, 60, 70, 80, 90, 95, 100]
        sizes = {}
        for q in candidates:
            sizes[q] = size_at_quality(img, width, q, tmp_path)

        for q, size in sizes.items():
            diff = abs(size - target_bytes)
            if best is None or diff < best[0]:
                best = (diff, size, q, width)

        # Refinar entre las dos calidades muestreadas mas cercanas al objetivo
        sorted_q = sorted(candidates, key=lambda q: sizes[q])
        for i in range(len(sorted_q) - 1):
            q_lo, q_hi = sorted_q[i], sorted_q[i + 1]
            lo, hi = min(q_lo, q_hi), max(q_lo, q_hi)
            if sizes[lo] <= target_bytes <= sizes[hi] or sizes[hi] <= target_bytes <= sizes[lo]:
                while hi - lo > 1:
                    mid = (lo + hi) // 2
                    size = size_at_quality(img, width, mid, tmp_path)
                    diff = abs(size - target_bytes)
                    if diff < best[0]:
                        best = (diff, size, mid, width)
                    if size < target_bytes:
                        lo = mid
                    else:
                        hi = mid
                break

        max_quality_size = sizes[100]
        if max_quality_size <= target_bytes + TOLERANCE:
            # Incluso a calidad maxima no llega (o cae dentro del margen): es lo mejor posible.
            return best[1], best[2], best[3]

        if best[1] <= target_bytes + TOLERANCE:
            return best[1], best[2], best[3]

        # Sigue pesando mas que el objetivo incluso ajustando calidad: reducir ancho y reintentar
        if width <= MIN_W:
            return best[1], best[2], best[3]
        width = max(MIN_W, int(width * 0.85))


processed = []
skipped_already_ok = 0
tmp_path = os.path.join(ROOT, "_tmp_convert.webp")

for dirpath, _dirnames, filenames in os.walk(ROOT):
    for fname in sorted(filenames):
        lower = fname.lower()
        if not lower.endswith(SOURCE_EXTS):
            continue
        path = os.path.join(dirpath, fname)
        if not os.path.isfile(path):
            continue

        orig_size = os.path.getsize(path)

        # Si ya es WebP y pesa <= 30KB (dentro del margen), no tocar nada.
        if lower.endswith(".webp") and orig_size <= TARGET_BYTES + TOLERANCE:
            skipped_already_ok += 1
            continue

        img = Image.open(path)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        if img.width > MAX_W:
            new_h = max(1, int(img.height * (MAX_W / img.width)))
            img = img.resize((MAX_W, new_h), Image.LANCZOS)

        final_size, quality, width = encode_to_target(img, TARGET_BYTES, tmp_path)

        base, ext = os.path.splitext(path)
        new_path = base + ".webp"
        os.replace(tmp_path, new_path)
        if ext.lower() != ".webp" and os.path.exists(path) and path != new_path:
            os.remove(path)

        processed.append((new_path, orig_size, final_size, quality, width))

if os.path.exists(tmp_path):
    os.remove(tmp_path)

total_orig = sum(r[1] for r in processed)
total_new = sum(r[2] for r in processed)
for path, o, n, q, w in processed:
    pct = (1 - n / o) * 100 if o else 0
    print(f"{path}: {o/1024:.0f}KB -> {n/1024:.1f}KB (q={q}, w={w}px, {pct:.0f}% menos)")

print(f"\nConvertidas/reducidas: {len(processed)} imagenes (objetivo ~{TARGET_BYTES/1024:.0f}KB)")
print(f"Ya estaban en WebP <= {(TARGET_BYTES + TOLERANCE)/1024:.0f}KB, sin tocar: {skipped_already_ok}")
if total_orig:
    print(f"TOTAL: {total_orig/1024:.0f}KB -> {total_new/1024:.0f}KB ({(1 - total_new/total_orig)*100:.0f}% menos)")

# --- Consolidacion de duplicados (siempre al final, con todo ya en .webp) ---
dup_deleted, dup_rescued = consolidate_duplicates(ROOT)
print(f"\nDuplicados eliminados: {len(dup_deleted)}")
for src in dup_deleted[:10]:
    print(f"  - borrado: {src}")
if len(dup_deleted) > 10:
    print(f"  ... y {len(dup_deleted) - 10} mas")
print(f"Duplicados rescatados (sin original, renombrados): {len(dup_rescued)}")
for src, dst in dup_rescued[:10]:
    print(f"  - {src} -> {dst}")
if len(dup_rescued) > 10:
    print(f"  ... y {len(dup_rescued) - 10} mas")
