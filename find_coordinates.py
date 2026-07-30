"""
find_coordinates.py
Only needed for template_mode: "image".
Overlays a numbered pixel grid on your certificate design so you can read off
the x,y coordinates to type into config.yaml (image_text_fields).

Usage:
    python find_coordinates.py
Output:
    output/grid_preview.png   <- open this, find where the name/date should go,
                                 read the nearest gridline numbers, put them in config.yaml
"""
import os
import yaml
from PIL import Image, ImageDraw, ImageFont
from pdf2image import convert_from_path


def load_base_image(template_path, dpi=300):
    if template_path.lower().endswith(".pdf"):
        pages = convert_from_path(template_path, dpi=dpi)
        return pages[0].convert("RGB")
    return Image.open(template_path).convert("RGB")


def main():
    with open("config.yaml") as f:
        cfg = yaml.safe_load(f)

    img = load_base_image(cfg["image_template_path"], dpi=cfg.get("image_render_dpi", 300))
    draw = ImageDraw.Draw(img)
    w, h = img.size
    step = 100

    for x in range(0, w, step):
        draw.line([(x, 0), (x, h)], fill=(255, 0, 0), width=1)
        draw.text((x + 2, 2), str(x), fill=(255, 0, 0))
    for y in range(0, h, step):
        draw.line([(0, y), (w, y)], fill=(0, 120, 255), width=1)
        draw.text((2, y + 2), str(y), fill=(0, 120, 255))

    os.makedirs(cfg.get("output_dir", "output"), exist_ok=True)
    out_path = os.path.join(cfg.get("output_dir", "output"), "grid_preview.png")
    img.save(out_path)
    print(f"Grid preview saved to {out_path} (image is {w}x{h}px). Open it, then update config.yaml.")


if __name__ == "__main__":
    main()
