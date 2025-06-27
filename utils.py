# utils.py
import os
from functools import lru_cache
from PIL import Image, ImageTk

PREVIEW_FOLDER = "previews"
SCHITE_FOLDER  = "schite"

@lru_cache(maxsize=64)
def make_preview(code, max_w=280, max_h=180):
    path = os.path.join(PREVIEW_FOLDER, f"{code}.png")
    if not os.path.exists(path):
        return None
    img_raw = Image.open(path)
    w, h = img_raw.size
    ratio = min(max_w / w, max_h / h, 1.0)
    new_size = (int(w * ratio), int(h * ratio))
    img_resized = img_raw.resize(new_size, Image.Resampling.LANCZOS)
    return ImageTk.PhotoImage(img_resized)

def get_schita_path(code):
    path = os.path.join(SCHITE_FOLDER, f"{code}.png")
    return path if os.path.exists(path) else None

