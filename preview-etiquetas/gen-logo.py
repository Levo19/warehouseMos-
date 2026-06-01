"""
Genera el bitmap del logo "Caserito Tony's" para la etiqueta.
- Tamaño target: 144 x 28 dots (cabe en el header arriba-izq de Layout 5)
- Casita SVG (líneas finas) + "Caserito" (Pacifico) + "TONY'S" (Lilita One)

Salidas:
  logo-preview.png       — preview visual del logo (a escala 1:1)
  logo-preview-4x.png    — preview escalado 4x (para ver bien en pantalla)
  logo-zpl-gf.txt        — ^GF block listo para ZPL (Labelary preview)
  logo-tspl-bitmap.txt   — bytes para TSPL2 BITMAP command (impresión real)
"""
from PIL import Image, ImageDraw, ImageFont
import os, sys

sys.stdout.reconfigure(encoding='utf-8')
W, H = 180, 36   # mas grande para mejor legibilidad
FONTS = 'fonts'
pacifico = ImageFont.truetype(os.path.join(FONTS, 'Pacifico-Regular.ttf'), 19)
lilita   = ImageFont.truetype(os.path.join(FONTS, 'LilitaOne-Regular.ttf'), 18)

img = Image.new('1', (W, H), 1)  # 1=blanco
d = ImageDraw.Draw(img)

# ── Casita mas grande (x=2..32, y=4..34) ──
d.line([(3, 17), (17, 4), (31, 17)], fill=0, width=3)   # techo (grueso)
d.rectangle([(7, 17), (27, 33)], outline=0, width=2)    # cuerpo
d.rectangle([(14, 24), (20, 33)], fill=0)               # puerta
d.rectangle([(22, 8), (26, 17)], fill=0)                # chimenea

# ── Texto a la derecha (x=38..180) ──
d.text((38, -3), 'Caserito', font=pacifico, fill=0)
d.text((38, 17), "TONY'S", font=lilita, fill=0)

img.save('logo-preview.png')
img.resize((W * 4, H * 4), Image.NEAREST).save('logo-preview-4x.png')
print(f'✓ logo-preview.png ({W}x{H})')
print(f'✓ logo-preview-4x.png ({W*4}x{H*4})')

# ── ZPL ^GFA — invertir bits ──
# PIL mode '1' bytes: 1=blanco, 0=negro. ZPL ^GFA: 1=ink, 0=no ink → INVERTIR
raw = img.tobytes()
bytes_per_row = (W + 7) // 8
total = bytes_per_row * H

# Asegurar tamaño correcto (pad si PIL devuelve más)
raw = raw[:total]
inverted = bytes(b ^ 0xFF for b in raw)
zpl_hex = inverted.hex().upper()
zpl_block = f'^GFA,{total},{total},{bytes_per_row},{zpl_hex}'
with open('logo-zpl-gf.txt', 'w') as f:
    f.write(zpl_block)
print(f'✓ logo-zpl-gf.txt ({total} bytes, {bytes_per_row} bpr)')

# ── TSPL2 BITMAP — bytes raw, NO invertir ──
# TSPL2: 0=ink, 1=no ink → mismo orden que PIL (NO invertir)
# Comando: BITMAP x, y, width_bytes, height, mode, "raw"
# Para script de GAS guardamos los bytes en hex (luego se convierten)
tspl_hex = raw.hex().upper()
tspl_block = f'BITMAP {{x}},{{y}},{bytes_per_row},{H},0,"<{total} bytes raw>"\n'
tspl_block += f'# bytes_per_row={bytes_per_row}, height={H}, total={total}\n'
tspl_block += f'# hex={tspl_hex}\n'
with open('logo-tspl-bitmap.txt', 'w') as f:
    f.write(tspl_block)
print(f'✓ logo-tspl-bitmap.txt ({total} bytes raw)')

print('')
print(f'📐 Dimensiones bitmap: {W}x{H} dots = {W/8:.1f}x{H/8:.1f} mm @ 8 dpmm')
