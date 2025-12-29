import os
from PIL import Image, ImageOps, ImageChops, ImageStat

MIN_RATIO = 0.5
SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}

# Heuristic sensitivity: lower = more likely to treat pixels as "content"
DIFF_THRESHOLD = 18  # 10~30 typical; increase if too sensitive


def script_folder() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def avg_corner_color(img: Image.Image, sample: int = 30):
    """
    Estimate background color by averaging 4 corner patches.
    Returns an RGB tuple.
    """
    im = img.convert("RGB")
    w, h = im.size
    s = min(sample, w, h)

    patches = [
        im.crop((0, 0, s, s)),
        im.crop((w - s, 0, w, s)),
        im.crop((0, h - s, s, h)),
        im.crop((w - s, h - s, w, h)),
    ]

    # Average each patch, then average the 4 results
    colors = []
    for p in patches:
        stat = ImageStat.Stat(p)
        colors.append(tuple(int(x) for x in stat.mean))

    r = sum(c[0] for c in colors) // len(colors)
    g = sum(c[1] for c in colors) // len(colors)
    b = sum(c[2] for c in colors) // len(colors)
    return (r, g, b)


def content_bbox(img: Image.Image, bg_rgb, diff_threshold: int = DIFF_THRESHOLD):
    """
    Return bounding box of non-background content.
    Heuristic: difference vs flat background + threshold.
    """
    im = img.convert("RGB")
    w, h = im.size
    bg = Image.new("RGB", (w, h), bg_rgb)

    diff = ImageChops.difference(im, bg).convert("L")
    # Threshold: pixels above threshold are considered content
    mask = diff.point(lambda p: 255 if p > diff_threshold else 0)

    return mask.getbbox()  # (left, upper, right, lower) or None


def risky_crop(img: Image.Image, crop_top: int, crop_bottom: int) -> bool:
    """
    Determine if cropping top/bottom will likely cut important content.
    """
    bg = avg_corner_color(img)
    bbox = content_bbox(img, bg)

    if bbox is None:
        # Almost all background; safe
        return False

    _, content_top, _, content_bottom = bbox
    h = img.size[1]
    safe_upper = crop_top
    safe_lower = h - crop_bottom

    # If content extends into areas that would be removed, it's risky
    if content_top < safe_upper or content_bottom > safe_lower:
        return True
    return False


def crop_in_place(img: Image.Image, crop_top: int, crop_bottom: int) -> Image.Image:
    w, h = img.size
    return img.crop((0, crop_top, w, h - crop_bottom))


def zoom_out_to_canvas(img: Image.Image, target_h: int) -> Image.Image:
    """
    Create (w, target_h) canvas and scale original down to fit, centered.
    """
    img_rgb = img.convert("RGB")
    w, h = img_rgb.size

    bg = avg_corner_color(img_rgb)
    canvas = Image.new("RGB", (w, target_h), bg)

    # scale to fit inside canvas
    scale = min(w / w, target_h / h)  # effectively target_h/h when too tall
    new_w = max(1, int(round(w * scale)))
    new_h = max(1, int(round(h * scale)))

    resized = img_rgb.resize((new_w, new_h), Image.Resampling.LANCZOS)

    # center paste
    x = (w - new_w) // 2
    y = (target_h - new_h) // 2
    canvas.paste(resized, (x, y))
    return canvas


def ask_user_choice(filename: str) -> str:
    """
    Popup dialog with 3 options:
    - crop
    - zoom
    - skip

    Returns: "crop" | "zoom" | "skip"
    """
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        msg = (
            f"Potential text/pattern cut detected:\n\n{filename}\n\n"
            "Choose an action:\n"
            "YES  = Crop anyway (may cut)\n"
            "NO   = Zoom-out to fit canvas (no cut)\n"
            "CANCEL = Skip this image"
        )
        res = messagebox.askyesnocancel("Crop Warning", msg)
        root.destroy()

        if res is True:
            return "crop"
        if res is False:
            return "zoom"
        return "skip"

    except Exception:
        # Fallback to console
        print(f"\nWARNING: Potential text/pattern cut: {filename}")
        print("1) Crop anyway  2) Zoom-out to fit  3) Skip")
        while True:
            x = input("Select 1/2/3: ").strip()
            if x == "1":
                return "crop"
            if x == "2":
                return "zoom"
            if x == "3":
                return "skip"


def save_overwrite(original_path: str, out_img: Image.Image, original_format: str):
    """
    Overwrite original file. Keep format when possible.
    """
    fmt = (original_format or "").upper()
    ext = os.path.splitext(original_path)[1].lower()

    # Ensure appropriate mode for JPEG
    if fmt in {"JPEG", "JPG"} or ext in {".jpg", ".jpeg"}:
        if out_img.mode not in ("RGB", "L"):
            out_img = out_img.convert("RGB")
        out_img.save(original_path, quality=95, optimize=True)
    else:
        # For PNG/WebP/TIFF etc.
        out_img.save(original_path)


def process_folder(base_folder: str):
    for name in os.listdir(base_folder):
        ext = os.path.splitext(name.lower())[1]
        if ext not in SUPPORTED_EXTS:
            continue

        path = os.path.join(base_folder, name)

        try:
            with Image.open(path) as im:
                original_format = im.format
                im = ImageOps.exif_transpose(im)

                w, h = im.size
                ratio = w / h

                if ratio >= MIN_RATIO:
                    print(f"OK:      {name} | {w}x{h} (r={ratio:.3f})")
                    continue

                target_h = int(round(w / MIN_RATIO))  # = 2*w when min_ratio=0.5
                excess = h - target_h
                crop_top = excess // 2
                crop_bottom = excess - crop_top

                # Risk check
                is_risky = risky_crop(im, crop_top, crop_bottom)

                if is_risky:
                    choice = ask_user_choice(name)
                else:
                    choice = "crop"

                if choice == "skip":
                    print(f"SKIP:    {name} | risky crop")
                    continue

                if choice == "crop":
                    out = crop_in_place(im, crop_top, crop_bottom)
                    save_overwrite(path, out, original_format)
                    nw, nh = out.size
                    print(f"CROPPED: {name} | {w}x{h} -> {nw}x{nh}")
                elif choice == "zoom":
                    out = zoom_out_to_canvas(im, target_h=target_h)
                    save_overwrite(path, out, original_format)
                    nw, nh = out.size
                    print(f"ZOOMFIT: {name} | {w}x{h} -> {nw}x{nh} (no cut)")

        except Exception as e:
            print(f"FAIL: {name} | {e}")


if __name__ == "__main__":
    BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))
    process_folder(BASE_FOLDER)
