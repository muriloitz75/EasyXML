from PIL import Image, ImageDraw, ImageFont
import os

# Cria uma imagem 256x256 com fundo transparente
icon_size = (256, 256)
icon = Image.new('RGBA', icon_size, (0, 0, 0, 0))
draw = ImageDraw.Draw(icon)

# Define cores
bg_color = (52, 152, 219)  # Azul
text_color = (255, 255, 255)  # Branco
xml_color = (231, 76, 60)  # Vermelho
excel_color = (46, 204, 113)  # Verde

# Desenha um círculo como fundo
center = (icon_size[0] // 2, icon_size[1] // 2)
radius = min(icon_size) // 2 - 10
draw.ellipse((center[0] - radius, center[1] - radius, 
              center[0] + radius, center[1] + radius), 
             fill=bg_color)

# Tenta carregar uma fonte, ou usa a fonte padrão
try:
    font = ImageFont.truetype("arial.ttf", 80)
    small_font = ImageFont.truetype("arial.ttf", 40)
except IOError:
    font = ImageFont.load_default()
    small_font = ImageFont.load_default()

# Desenha o texto "XML" na parte superior
text = "XML"
text_width = draw.textlength(text, font=font)
text_position = (center[0] - text_width // 2, center[1] - 60)
draw.text(text_position, text, font=font, fill=xml_color)

# Desenha o texto "Excel" na parte inferior
text = "Excel"
text_width = draw.textlength(text, font=small_font)
text_position = (center[0] - text_width // 2, center[1] + 20)
draw.text(text_position, text, font=small_font, fill=excel_color)

# Salva como PNG e ICO
icon.save("icon.png")
icon.save("icon.ico", format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])

print("Ícone criado com sucesso!")
