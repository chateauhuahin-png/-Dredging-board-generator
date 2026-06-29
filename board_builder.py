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
                continue
            for kw in PHOTO_LABELS:
                if kw in shape.text and kw not in label_x:
                    label_x[kw] = (shape.left or 0) + (shape.width or 0) // 2
        imgs = []
        for shape in photo_slide.shapes:
            if shape.shape_type == 13:
                cx = (shape.left or 0) + (shape.width or 0) // 2
                imgs.append((cx, shape.image.blob, shape.image.ext))
        imgs.sort(key=lambda x: x[0])
        if imgs:
            res = {}
            for kw, lx in label_x.items():
                cl = min(imgs, key=lambda img: abs(img[0] - lx))
                fname = os.path.join(work_dir, f"photo_{kw}.{cl[2]}")
                with open(fname, "wb") as f:
                    f.write(cl[1])
                res[kw] = fname
            photo_before = res.get("ก่อน")
            photo_during = res.get("ระหว่าง")
            photo_after  = res.get("หลัง")
        del imgs

    # ── extract media from key slides (sorted by X position = left→right) ──
    os.makedirs(med_dir, exist_ok=True)
    def _extract(shapes, prefix):
        # Sort by left position so filenames reflect left→right order
        pic_shapes = []
        for shape in shapes:
            if shape.shape_type == 13:
                pic_shapes.append(shape)
            elif shape.shape_type == 6:
                _extract(shape.shapes, f"{prefix}g")
        pic_shapes.sort(key=lambda s: (s.left or 0))
        for j, shape in enumerate(pic_shapes):
            img = shape.image
            fname = f"{prefix}_{j:03d}.{img.ext}"
            with open(os.path.join(med_dir, fname), "wb") as f:
                f.write(img.blob)

    key_slides = list(set(s for s in [cfg["map"], cfg["letter1"], cfg["letter2"],
                                       cfg["surv"], cfg["des"], cfg["cross"],
                                       cfg["vol"], cfg["boq"]] if s))
    for si in key_slides:
        _extract(prs.slides[si-1].shapes, f"s{si:02d}")

    # ── PIL composite for map slide (while PPTX still open) ──────────────
    import io as _io
    map_composite = None
    if cfg["map"]:
        try:
            slide  = prs.slides[cfg["map"] - 1]
            sw     = prs.slide_width
            sh_emu = prs.slide_height
            OUT_W  = 2400
            sc     = OUT_W / sw
            OUT_H  = int(sh_emu * sc)
            base   = Image.new("RGB", (OUT_W, OUT_H), (255, 255, 255))

            def _paste_pics(shapes, ox=0, oy=0, sx=1.0, sy=1.0):
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
                            px_, py_ = int(ax * sc), int(ay * sc)
                            pw,  ph  = int(aw * sc), int(ah * sc)
                            im = Image.open(_io.BytesIO(s.image.blob)).convert("RGBA")
                            im = im.resize((pw, ph), Image.LANCZOS)
                            bg = Image.new("RGB", (pw, ph), (255, 255, 255))
                            bg.paste(im, mask=im.split()[3])
                            base.paste(bg, (px_, py_))
                            im.close(); bg.close(); del im, bg
                        except Exception as e:
                            print(f"    map pic: {e}")
                    elif s.shape_type == 6:
                        _paste_pics(s.shapes, ax, ay, sx, sy)

            _paste_pics(slide.shapes)
            map_composite = os.path.join(work_dir, "map_composite.jpg")
            base.save(map_composite, "JPEG", quality=95)
            base.close(); del base
            gc.collect()
            print("  Map PIL composite saved.")
        except Exception as e:
            print(f"  Map composite error: {e}")
            map_composite = None

    # ── close PPTX now — free all memory before building board ───────────
    del prs
    gc.collect()
    print("PPTX closed, memory freed.")

    return cfg, title1, title2, agency, photo_before, photo_during, photo_after, map_composite


def build_board(pptx_path, work_dir, output_path):
    """Main function: build board from PPTX"""
    logo_path = os.path.join(BASE_DIR, "fonts", "logo.png")
    os.makedirs(work_dir, exist_ok=True)
    med_dir = os.path.join(work_dir, "media")

    # ── 1. Read ALL data from PPTX in one pass, then close it ────────────
    cfg, title1, title2, agency, photo_before, photo_during, photo_after, map_composite = \
        _parse_pptx_once(pptx_path, work_dir, med_dir)

    # 3. Find image files for each section
    def find_img(si):
        """Find best (largest) image for slide si"""
        if si is None:
            return None
        candidates = []
        for f in sorted(os.listdir(med_dir)):
            if f.startswith(f"s{si:02d}_"):
                p = os.path.join(med_dir, f)
                try:
                    img = Image.open(p)
                    candidates.append((p, img.size[0]*img.size[1]))
                    img.close()
                except Exception:
                    pass
        if not candidates:
            return None
        candidates.sort(key=lambda x: x[1], reverse=True)
        return candidates[0][0]

    def find_all_imgs(si):
        """Find ALL images for slide si sorted by area (largest first)"""
        if si is None:
            return []
        candidates = []
        for f in sorted(os.listdir(med_dir)):
            if f.startswith(f"s{si:02d}_"):
                p = os.path.join(med_dir, f)
                try:
                    img = Image.open(p)
                    candidates.append((p, img.size[0]*img.size[1]))
                    img.close()
                except Exception:
                    pass
        candidates.sort(key=lambda x: x[1], reverse=True)
        return [p for p, _ in candidates]

    def find_boq_imgs(si):
        """BOQ slide has 2 images: tall one = ปร.6, wide one = ปร.4"""
        if si is None:
            return None, None
        candidates = []
        for f in sorted(os.listdir(med_dir)):
            if f.startswith(f"s{si:02d}_"):
                p = os.path.join(med_dir, f)
                try:
                    img = Image.open(p)
                    candidates.append((p, img.size[0], img.size[1]))
                except Exception:
                    pass
        if not candidates:
            return None, None
        if len(candidates) < 2:
            return candidates[0][0], None
        tall = sorted(candidates, key=lambda x: x[2]/max(x[1],1), reverse=True)
        wide = sorted(candidates, key=lambda x: x[1]/max(x[2],1), reverse=True)
        return tall[0][0], wide[0][0]

    # ── 2. Render map slide: Graph API → map_composite → find_img ────────
    print("Rendering map slide...")
    map_jpg = None
    if cfg["map"]:
        # Try Graph API first (perfect render with text/lines)
        try:
            from graph_convert import slide_to_png, is_configured
            if is_configured():
                print("  Using Microsoft Graph API...")
                png = slide_to_png(pptx_path, cfg["map"], work_dir)
                if png:
                    map_jpg = os.path.join(work_dir, "map_slide.jpg")
                    img = Image.open(png).convert("RGB")
                    img.save(map_jpg, "JPEG", quality=95)
                    img.close()
                    print("  Graph API: success")
        except Exception as e:
            print(f"  Graph API error: {e}")

        if not map_jpg:
            # Use PIL composite (built during PPTX parse, no extra memory)
            map_jpg = map_composite
            if map_jpg:
                print("  Using PIL composite fallback")
            else:
                map_jpg = find_img(cfg["map"])
                print("  Using find_img fallback")

    print(f"Title: {title1} / {title2} / Agency: {agency}")

    # ── 3. Logo ──────────────────────────────────────────────────────────

    boq_img, price_img = find_boq_imgs(cfg["boq"])
    surv_imgs  = find_all_imgs(cfg["surv"])    # ← หลายรูป
    des_img    = find_img(cfg["des"])
    cross_imgs = find_all_imgs(cfg["cross"])   # ← หลายรูป
    vol_img    = find_img(cfg["vol"])
    lett2_imgs = find_all_imgs(cfg["letter2"]) if cfg["letter2"] else []  # ← หลายรูป

    # 7. Build board
    board = Image.new("RGB", (W, H), NAVY)
    draw  = ImageDraw.Draw(board)

    # Header
    draw.rectangle([MG, MG, W-MG, MG+HDR], fill=NAVY)
    logo_sz = 680

    # Logo — มุมบนซ้าย (paste with transparency, cap height to HDR)
    if logo_path and os.path.exists(logo_path):
        try:
            lg_img = Image.open(logo_path).convert("RGBA")
            lw, lh_img = lg_img.size
            sc = min(logo_sz / lw, (HDR - 20) / lh_img)
            nw, nh = int(lw * sc), int(lh_img * sc)
            lg_img = lg_img.resize((nw, nh), Image.LANCZOS)
            lx = XL + (logo_sz - nw) // 2
            ly_logo = MG + (HDR - nh) // 2
            board.paste(lg_img, (lx, ly_logo), mask=lg_img.split()[3])
        except Exception as e:
            print(f"  logo error: {e}")

    # ชื่อหน่วยงาน — มุมบนขวา (สีขาว บรรทัดเดียว ฟอนต์ใหญ่)
    agency_text = agency if agency else ""
    fa = fnt(122, bold=True)
    bb_a = draw.textbbox((0,0), agency_text, font=fa)
    aw = bb_a[2] - bb_a[0]
    ah = bb_a[3] - bb_a[1]
    right_x = W - MG - logo_sz
    ax = right_x - aw
    ay = MG + (HDR - ah) // 2
    draw.text((ax, ay), agency_text, font=fa, fill=WHITE)

    t1sz = 148 if len(title1) < 60 else 132
    f1 = fnt(t1sz, True); f2 = fnt(114)
    cx = W // 2
    bb1 = draw.textbbox((0,0), title1, font=f1)
    bb2 = draw.textbbox((0,0), title2, font=f2)
    draw.text((cx-(bb1[2]-bb1[0])//2, MG+45),  title1, font=f1, fill=GOLD)
    draw.text((cx-(bb2[2]-bb2[0])//2, MG+205), title2, font=f2, fill=WHITE)

    # Left column
    boq_h = int(CONH * 0.54)
    sec(draw, board, "ประมาณการ (ปร.6)", boq_img,   XL, CONY, CL, boq_h)
    sec(draw, board, "ประมาณการ (ปร.4)", price_img, XL, CONY+boq_h+GAP, CL, CONH-boq_h-GAP)

    # Middle column
    map_h = int(CONH * 0.50)
    sec(draw, board, "แผนที่และจุดดำเนินการ (มาตราส่วน 1:50,000)",
        map_jpg, XM, CONY, CM, map_h, lsz=68)
    sy = CONY + map_h + GAP
    sh = int(CONH * 0.265)
    sec_multi(draw, board, "ตารางการสำรวจ",    surv_imgs, XM, sy,          CM, sh)
    sec(draw, board, "ตารางการออกแบบ",         des_img,   XM, sy+sh+GAP,   CM, CONH-map_h-GAP-sh-GAP)

    # Right column
    ch = int(CONH * 0.375)
    sec_multi(draw, board, "รูปตัดตามขวางขุดลอกลำน้ำ", cross_imgs, XR, CONY, CR, ch)

    ly = CONY + ch + GAP
    lh = int(CONH * 0.37)
    sec_multi(draw, board, "หนังสือตรวจสอบความซ้ำซ้อน", lett2_imgs, XR, ly, CR, lh, lsz=66)

    vy = ly + lh + GAP
    sec(draw, board, "ตารางคำนวณปริมาตรดินตะกอน", vol_img, XR, vy, CR, CONH-ch-GAP-lh-GAP)

    # Photo strip
    phy = CONY + CONH + GAP
    PW  = (UW - 2*GAP) // 3
    for idx, (lbl, ph) in enumerate(zip(
        ["ภาพก่อนปฏิบัติงาน", "ภาพระหว่างปฏิบัติงาน", "ภาพหลังปฏิบัติงาน"],
        [photo_before, photo_during, photo_after]
    )):
        px = XL + idx * (PW + GAP)
        draw.rectangle([px, phy, px+PW, phy+LH], fill=NAVY)
        fp = fnt(84, True)
        bbl = draw.textbbox((0,0), lbl, font=fp)
        draw.text((px+(PW-(bbl[2]-bbl[0]))//2, phy+(LH-(bbl[3]-bbl[1]))//2),
                  lbl, font=fp, fill=GOLD)
        if ph and os.path.exists(ph):
            board.paste(fit(ph, PW, PHH-LH), (px, phy+LH))

    board.save(output_path, "JPEG", quality=95, dpi=(150, 150))
    print(f"Saved: {output_path}")
    return output_path
