"""
Board Builder - สร้างบอร์ดชี้แจงจาก PPTX + รูปภาพ
"""
import os, copy, gc
from PIL import Image, ImageDraw, ImageFont, ImageFile
from pptx import Presentation

ImageFile.LOAD_TRUNCATED_IMAGES = True

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
FONT_BOLD = os.path.join(BASE_DIR, "fonts", "THSarabun Bold.ttf")
FONT_REG  = os.path.join(BASE_DIR, "fonts", "THSarabun.ttf")

# 150 DPI @ 120×80 cm
W, H   = 7087, 4724
NAVY   = (11, 20, 100)
GOLD   = (255, 215, 0)
WHITE  = (255, 255, 255)
MG, GAP = 120, 36
UW = W - 2*MG
UH = H - 2*MG
HDR  = 400
PHH  = 840
LH   = 100
CL   = int(UW * 0.295)
CM   = int(UW * 0.325)
CR   = UW - CL - CM - 2*GAP
CONH = UH - HDR - GAP - PHH - GAP
CONY = MG + HDR + GAP
XL   = MG
XM   = XL + CL + GAP
XR   = XM + CM + GAP


def fnt(sz, bold=False):
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REG, sz)


def fit(path, w, h, bg=WHITE):
    canvas = Image.new("RGB", (w, h), bg)
    try:
        img = Image.open(path).convert("RGB")
        iw, ih = img.size
        sc = min(w / iw, h / ih)
        nw, nh = int(iw * sc), int(ih * sc)
        img = img.resize((nw, nh), Image.LANCZOS)
        canvas.paste(img, ((w - nw) // 2, (h - nh) // 2))
        img.close()
        del img
    except Exception as e:
        print(f"  fit error {path}: {e}")
    return canvas


def sec(draw, board, label, img_path, x, y, w, h, lsz=76):
    draw.rectangle([x, y, x+w, y+LH], fill=NAVY)
    f = fnt(lsz, bold=True)
    bb = draw.textbbox((0, 0), label, font=f)
    draw.text((x + (w-(bb[2]-bb[0]))//2, y + (LH-(bb[3]-bb[1]))//2),
              label, font=f, fill=GOLD)
    if img_path and os.path.exists(img_path):
        tile = fit(img_path, w, h - LH)
        board.paste(tile, (x, y + LH))
        tile.close()
        del tile
        gc.collect()
    else:
        draw.rectangle([x, y+LH, x+w, y+h], fill=WHITE)


def sec_multi(draw, board, label, img_paths, x, y, w, h, lsz=76):
    """Section with multiple images arranged left to right"""
    draw.rectangle([x, y, x+w, y+LH], fill=NAVY)
    f = fnt(lsz, bold=True)
    bb = draw.textbbox((0, 0), label, font=f)
    draw.text((x + (w-(bb[2]-bb[0]))//2, y + (LH-(bb[3]-bb[1]))//2),
              label, font=f, fill=GOLD)
    valid = [p for p in img_paths if p and os.path.exists(p)]
    if not valid:
        draw.rectangle([x, y+LH, x+w, y+h], fill=WHITE)
        return
    n = len(valid)
    slot_w = w // n
    img_h  = h - LH
    for i, p in enumerate(valid):
        tile = fit(p, slot_w, img_h)
        board.paste(tile, (x + i * slot_w, y + LH))
        tile.close()
        del tile
    gc.collect()


def extract_media(pptx_path, out_dir, slide_indices):
    """Extract images from specific slides (1-based index)"""
    os.makedirs(out_dir, exist_ok=True)
    prs = Presentation(pptx_path)

    def _extract(shapes, prefix):
        for j, shape in enumerate(shapes):
            if shape.shape_type == 13:
                img = shape.image
                fname = f"{prefix}_{j}.{img.ext}"
                with open(os.path.join(out_dir, fname), "wb") as f:
                    f.write(img.blob)
            elif shape.shape_type == 6:
                _extract(shape.shapes, f"{prefix}g{j}")

    for si in slide_indices:
        _extract(prs.slides[si-1].shapes, f"s{si:02d}")

    return prs


def extract_single_slide(pptx_path, slide_idx, work_dir):
    """Extract one slide into its own PPTX to reduce memory usage"""
    prs_full = Presentation(pptx_path)
    slide    = prs_full.slides[slide_idx - 1]
    sw, sh   = prs_full.slide_width, prs_full.slide_height

    prs_new = Presentation()
    prs_new.slide_width  = sw
    prs_new.slide_height = sh
    if len(prs_new.slides._sldIdLst) > 0:
        prs_new.slides._sldIdLst.remove(prs_new.slides._sldIdLst[0])

    new_slide = prs_new.slides.add_slide(prs_new.slide_layouts[6])
    sp_tree = new_slide.shapes._spTree
    for child in list(sp_tree):
        sp_tree.remove(child)
    for child in slide.shapes._spTree:
        sp_tree.append(copy.deepcopy(child))
    for rel in slide.part.rels.values():
        if "image" in rel.reltype:
            try:
                new_slide.part.relate_to(rel.target_part, rel.reltype)
            except Exception:
                pass

    out = os.path.join(work_dir, "map_single.pptx")
    prs_new.save(out)
    del prs_full
    return out, sw, sh


def render_map_slide(pptx_path, slide_idx, work_dir):
    """
    Render map slide to JPG.
    Priority 1: Microsoft Graph API (perfect quality)
    Priority 2: PIL composite — paste all images at correct EMU positions (no text/shapes)
    """
    import io as _io

    # ── Priority 1: Graph API ─────────────────────────────────────────────
    try:
        from graph_convert import slide_to_png, is_configured
        if is_configured():
            print("  Using Microsoft Graph API...")
            png = slide_to_png(pptx_path, slide_idx, work_dir)
            if png:
                jpg = os.path.join(work_dir, "map_slide.jpg")
                Image.open(png).convert("RGB").save(jpg, "JPEG", quality=95)
                print("  Graph API: success")
                return jpg
            print("  Graph API: no output — falling back to PIL composite")
    except Exception as e:
        print(f"  Graph API error: {e} — falling back to PIL composite")

    # ── Priority 2: PIL composite (images only, correct positions) ────────
    try:
        print("  PIL composite render...")
        prs  = Presentation(pptx_path)
        slide = prs.slides[slide_idx - 1]
        sw   = prs.slide_width
        sh   = prs.slide_height
        OUT_W = 2400
        scale = OUT_W / sw
        OUT_H = int(sh * scale)
        base  = Image.new("RGB", (OUT_W, OUT_H), (255, 255, 255))

        def paste_pics(shapes, ox=0, oy=0, sx=1.0, sy=1.0):
            for s in shapes:
                try:
                    ax = ox + (s.left  or 0) * sx
                    ay = oy + (s.top   or 0) * sy
                    aw = (s.width  or 0) * sx
                    ah = (s.height or 0) * sy
                except Exception:
                    continue
                if s.shape_type == 13 and aw > 0 and ah > 0:
                    try:
                        px_, py_ = int(ax * scale), int(ay * scale)
                        pw,  ph  = int(aw * scale), int(ah * scale)
                        im = Image.open(_io.BytesIO(s.image.blob)).convert("RGBA")
                        im = im.resize((pw, ph), Image.LANCZOS)
                        bg = Image.new("RGB", (pw, ph), (255, 255, 255))
                        bg.paste(im, mask=im.split()[3])
                        base.paste(bg, (px_, py_))
                        im.close(); bg.close()
                        del im, bg
                    except Exception as e:
                        print(f"    pic paste: {e}")
                elif s.shape_type == 6:
                    paste_pics(s.shapes,
                               ox + (s.left or 0) * sx,
                               oy + (s.top  or 0) * sy,
                               sx, sy)

        paste_pics(slide.shapes)
        del prs
        gc.collect()

        jpg = os.path.join(work_dir, "map_slide.jpg")
        base.save(jpg, "JPEG", quality=95)
        base.close()
        print("  PIL composite: success")
        return jpg
    except Exception as e:
        print(f"  PIL composite error: {e}")
        return None


def get_title_from_pptx(pptx_path):
    """Extract title, location, and agency name from slide 1"""
    prs = Presentation(pptx_path)
    slide = prs.slides[0]
    title1, title2, agency = "งานขุดลอกลำน้ำ", "", ""

    for shape in slide.shapes:
        if not hasattr(shape, "text") or not shape.text.strip():
            continue
        text = shape.text.strip()

        # ดึงชื่อหน่วยงาน — หาจาก shape ที่มีคำว่า นพค หรือ นทพ
        if not agency and ("นพค" in text or "นทพ" in text):
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            # เอาเฉพาะบรรทัดที่เป็นชื่อหน่วยงาน
            for line in lines:
                if "นพค" in line or "นทพ" in line:
                    agency = line
                    break

        # ดึงชื่องานและที่ตั้ง
        if "งานขุด" in text:
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            title1 = f"{lines[0]} {lines[1]}" if len(lines) > 1 else lines[0]
            title2 = lines[2] if len(lines) > 2 else ""

    return title1, title2, agency


def detect_slide_map(pptx_path):
    """Auto-detect which slide numbers contain each section (slide 1 fixed, rest by keyword)"""
    prs = Presentation(pptx_path)
    result = {"map": None, "surv": None, "des": None, "cross": None,
              "vol": None, "boq": None, "letter1": None, "letter2": None}

    keywords = {
        "50,000":                       "map",
        "หนังสือขอรับการสนับสนุน":      "letter1",
        "ซ้ำซ้อน":                       "letter2",
        "ตารางการสำรวจ":                "surv",
        "ตารางการออกแบบ":               "des",
        "รูปตัดตามขวาง":                "cross",
        "ตารางคำนวณปริมาตร":            "vol",
        "แบบสรุปราคา":                  "boq",
    }

    for i, slide in enumerate(prs.slides, 1):
        if i == 1:
            continue  # สไลด์ 1 คือหน้าปก ข้ามเสมอ
        text = " ".join(s.text for s in slide.shapes if hasattr(s, "text"))
        for kw, key in keywords.items():
            if kw in text and result[key] is None:
                result[key] = i

    return result


def extract_photos_from_pptx(pptx_path, work_dir):
    """
    หาสไลด์ที่มี text label ก่อน/ระหว่าง/หลัง แล้ว match รูปภาพตามตำแหน่ง X
    Returns: (before_path, during_path, after_path)
    """
    prs = Presentation(pptx_path)
    LABELS = {"ก่อน": None, "ระหว่าง": None, "หลัง": None}

    # หาสไลด์ที่มีทั้ง 3 คำ
    photo_slide = None
    for slide in prs.slides:
        text_all = " ".join(s.text for s in slide.shapes if hasattr(s, "text"))
        if all(kw in text_all for kw in LABELS):
            photo_slide = slide
            break

    if not photo_slide:
        print("  Photo slide not found — no photos")
        return None, None, None

    # หาตำแหน่ง X กึ่งกลางของแต่ละ label
    label_x = {}
    for shape in photo_slide.shapes:
        if not hasattr(shape, "text"):
            continue
        for kw in LABELS:
            if kw in shape.text and kw not in label_x:
                label_x[kw] = (shape.left or 0) + (shape.width or 0) // 2

    # หารูปภาพทั้งหมดในสไลด์ พร้อม X กึ่งกลาง
    images = []
    for shape in photo_slide.shapes:
        if shape.shape_type == 13:
            cx = (shape.left or 0) + (shape.width or 0) // 2
            images.append((cx, shape.image.blob, shape.image.ext))
    images.sort(key=lambda x: x[0])

    if not images:
        print("  No images found in photo slide")
        return None, None, None

    # Match: label กับ image ที่ใกล้ที่สุด (ตาม X)
    result = {}
    for kw, lx in label_x.items():
        closest = min(images, key=lambda img: abs(img[0] - lx))
        fname = os.path.join(work_dir, f"photo_{kw}.{closest[2]}")
        with open(fname, "wb") as f:
            f.write(closest[1])
        result[kw] = fname
        print(f"  Photo '{kw}' → {fname}")

    return result.get("ก่อน"), result.get("ระหว่าง"), result.get("หลัง")


def _parse_pptx_once(pptx_path, work_dir, med_dir):
    """
    Open PPTX exactly once and extract ALL needed data:
    - slide map (cfg)
    - title/agency
    - photos (before/during/after)
    - media images for each section
    Then close PPTX and gc.collect() before heavy image work.
    """
    print("Opening PPTX (single pass)...")
    prs = Presentation(pptx_path)

    # ── slide map ────────────────────────────────────────────────────────────
    cfg = {"map": None, "surv": None, "des": None, "cross": None,
           "vol": None, "boq": None, "letter1": None, "letter2": None}
    keywords = {
        "50,000": "map", "หนังสือขอรับการสนับสนุน": "letter1",
        "ซ้ำซ้อน": "letter2", "ตารางการสำรวจ": "surv",
        "ตารางการออกแบบ": "des", "รูปตัดตามขวาง": "cross",
        "ตารางคำนวณปริมาตร": "vol", "แบบสรุปราคา": "boq",
    }

    # ── title/agency from slide 1 ─────────────────────────────────────────
    title1, title2, agency = "งานขุดลอกลำน้ำ", "", ""
    slide0 = prs.slides[0]
    for shape in slide0.shapes:
        if not hasattr(shape, "text") or not shape.text.strip():
            continue
        text = shape.text.strip()
        if not agency and ("นพค" in text or "นทพ" in text):
            for line in [l.strip() for l in text.split("\n") if l.strip()]:
                if "นพค" in line or "นทพ" in line:
                    agency = line; break
        if "งานขุด" in text:
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            title1 = f"{lines[0]} {lines[1]}" if len(lines) > 1 else lines[0]
            title2 = lines[2] if len(lines) > 2 else ""

    # ── slide scan (slides 2+) ────────────────────────────────────────────
    PHOTO_LABELS = {"ก่อน", "ระหว่าง", "หลัง"}
    photo_slide = None
    for i, slide in enumerate(prs.slides, 1):
        if i == 1:
            continue
        text_all = " ".join(s.text for s in slide.shapes if hasattr(s, "text"))
        for kw, key in keywords.items():
            if kw in text_all and cfg[key] is None:
                cfg[key] = i
        if photo_slide is None and all(kw in text_all for kw in PHOTO_LABELS):
            photo_slide = slide

    print(f"Slide map: {cfg}")

    # ── extract photos ────────────────────────────────────────────────────
    photo_before = photo_during = photo_after = None
    if photo_slide:
        label_x = {}
        for shape in photo_slide.shapes:
            if not hasattr(shape, "text"):
      