from PIL import Image

# Load the base image
img = Image.open("ai_brain_base.png").convert("RGBA")

# Desired icon sizes
sizes = [16, 32, 80, 128]

# Resize and save each icon
for size in sizes:
    resized = img.resize((size, size), Image.LANCZOS)
    resized.save(f"icon-{size}.png")
    print(f"Saved icon-{size}.png")