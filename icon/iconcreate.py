from PIL import Image

# Load the image
img_path = '/Users/mcatoglu/vsCode_projects/defineReplacer/icon_definefiller.png'
img = Image.open(img_path)

# Save as .ico format
icon_path = '/Users/mcatoglu/vsCode_projects/defineReplacer/icon_definefiller.ico'
img.save(icon_path, format='ICO', sizes=[(32, 32), (64, 64), (128, 128), (256, 256)])
