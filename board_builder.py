"""
Board Builder - สร้างบอร์ดชี้แจงจาก PPTX + รูปภาพ
"""
import os, subprocess, copy
from PIL import Image, ImageDraw, ImageFont, ImageFile
from pptx import Presentation

ImageFile.LOAD_TRUNCATED_IMAGES = True

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
FONT_BOLD = os.path.join(BASE_DIR, "fonts", "THSarabun Bold.ttf")
FONT_REG  = os.path.join(BASE_DIR, "fonts", "THSarabun.ttf")

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
    except Exception as e:
        print(f"  fit error {path}: {e}")
    return canvas


def sec(draw, board, label, img_path, x, y, w, h, lsz=68):
    draw.rectangle([x, y, x+w, y+LH], fill=NAVY)
    f = fnt(lsz, bold=True)
    bb = draw.textbbox((0, 0), label, font=f)
    draw.text((x + (w-(bb[2]-bb[0]))//2, y + (LH-(bb[3]-bb[1]))//2),
              label, font=f, fill=GOLD)
    if img_path and os.path.exists(img_path):
        board.paste(fit(img_path, w, h - LH), (x, y + LH))
    else:
        draw.rectangle([x, y+LH, x+w, y+h], fill=WHITE)


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
    Priority 1: Microsoft Graph API (PowerPoint renderer — perfect quality)
    Priority 2: LibreOffice composite render (fallback)
    """
    import io as _io

    # ── Priority 1: Microsoft Graph API ──────────────────────────────────────
    try:
        from graph_convert import slide_to_png, is_configured
        if is_configured():
            print("  Using Microsoft Graph API...")
            png = slide_to_png(pptx_path, slide_idx, work_dir)
            if png:
                jpg = os.path.join(work_dir, "map_slide.jpg")
                Image.open(png).convert("RGB").save(jpg, "JPEG", quality=92)
                print("  Graph API: success")
                return jpg
            print("  Graph API: no output, falling back to LibreOffice")
    except Exception as e:
        print(f"  Graph API error: {e}, falling back to LibreOffice")

    # ── Priority 2: LibreOffice composite render ──────────────────────────────
    single_pptx, sw, sh = extract_single_slide(pptx_path, slide_idx, work_dir)

    OUT_W = 1800
    scale = OUT_W / sw
    OUT_H = int(sh * scale)

    # ── Step 1: PPTX → ODP → PNG (ODP เป็น format ของ LibreOffice เอง render ได้ดีกว่า) ──
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "odp", single_pptx, "--outdir", work_dir],
        capture_output=True, timeout=60
    )
    odp_path = os.path.join(work_dir, "map_single.odp")
    render_src = odp_path if os.path.exists(odp_path) else single_pptx
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "png", render_src, "--outdir", work_dir],
        capture_output=True, timeout=60
    )
    lo_png = os.path.join(work_dir, "map_single.png")
    if os.path.exists(lo_png):
        base = Image.open(lo_png).convert("RGB").resize((OUT_W, OUT_H), Image.LANCZOS)
    else:
        base = Image.new("RGB", (OUT_W, OUT_H), (255, 255, 255))

    # ── Step 2: PIL pastes images at correct EMU positions on top ──────────
    prs   = Presentation(single_pptx)
    slide = prs.slides[0]

    def grp_transform(grp):
        try:
            xfrm = grp._element.grpSpPr.xfrm
            co, ce = xfrm.chOff, xfrm.chExt
            cox = co.x if co else 0;  coy = co.y if co else 0
            cex = ce.cx if ce else (grp.width or 1)
            cey = ce.cy if ce else (grp.height or 1)
            return cox, coy, (grp.width or cex)/cex, (grp.height or cey)/cey
        except Exception:
            return 0, 0, 1.0, 1.0

    def paste_pics(shapes, ox=0, oy=0, sx=1.0, sy=1.0, cox=0, coy=0):
        for s in shapes:
            try:
                ax = ox + ((s.left  or 0) - cox) * sx
                ay = oy + ((s.top   or 0) - coy) * sy
                aw = (s.width  or 0) * sx
                ah = (s.height or 0) * sy
            except Exception:
                continue
            px_, py_ = int(ax*scale), int(ay*scale)
            pw,  ph  = int(aw*scale), int(ah*scale)
            if s.shape_type == 13 and pw > 0 and ph > 0:
                try:
                    im = Image.open(_io.BytesIO(s.image.blob)).convert("RGBA")
                    im = im.resize((pw, ph), Image.LANCZOS)
                    bg = Image.new("RGB", (pw, ph), (255, 255, 255))
                    bg.paste(im, mask=im.split()[3])
                    base.paste(bg, (px_, py_))
                    del im, bg
                except Exception as e:
                    print(f"  pic: {e}")
            elif s.shape_type == 6:
                c_cox, c_coy, gsx, gsy = grp_transform(s)
                paste_pics(s.shapes, ax, ay, sx*gsx, sy*gsy, c_cox, c_coy)

    paste_pics(slide.shapes)
    del prs

    jpg = os.path.join(work_dir, "map_slide.jpg")
    base.save(jpg, "JPEG", quality=92)
    return jpg


def get_title_from_pptx(pptx_path):
    """Extract title and location from slide 1"""
    prs = Presentation(pptx_path)
    for shape in prs.slides[0].shapes:
        if hasattr(shape, "text") and "งานขุด" in shape.text:
            lines = [l.strip() for l in shape.text.strip().split("\n") if l.strip()]
            title1 = f"{lines[0]} {lines[1]}" if len(lines) > 1 else lines[0]
            title2 = lines[2] if len(lines) > 2 else ""
            return title1, title2
    return "งานขุดลอกลำน้ำ", ""


def detect_slide_map(pptx_path):
    """Auto-detect which slide numbers contain each section"""
    prs = Presentation(pptx_path)
    result = {"map": 4, "surv": None, "des": None, "cross": None,
              "vol": None, "boq": None, "letter1": 2, "letter2": 3}

    keywords = {
        "ตารางการสำรวจ": "surv",
        "ตารางการออกแบบ": "des",
        "รูปตัดตามขวาง": "cross",
        "ตารางคำนวณปริมาตร": "vol",
        "แบบสรุปราคา": "boq",
    }

    for i, slide in enumerate(prs.slides, 1):
        text = " ".join(s.text for s in slide.shapes if hasattr(s, "text"))
        for kw, key in keywords.items():
            if kw in text and result[key] is None:
                result[key] = i

    return result


def build_board(pptx_path, photo_before, photo_during, photo_after,
                work_dir, output_path, logo_path=None, map_override=None):
    # Always use built-in logo
    logo_path = os.path.join(BASE_DIR, "fonts", "logo.png")
    """Main function: build board from PPTX + 3 photos"""

    os.makedirs(work_dir, exist_ok=True)
    med_dir = os.path.join(work_dir, "media")

    # 1. Detect slides
    cfg = detect_slide_map(pptx_path)
    print(f"Slide map: {cfg}")

    # 2. Extract media from key slides
    key_slides = list(set([2, 3, cfg["map"], cfg["surv"], cfg["des"],
                            cfg["cross"], cfg["vol"], cfg["boq"]] if None not in
                           [cfg["surv"], cfg["des"], cfg["cross"], cfg["vol"], cfg["boq"]]
                           else [2, 3, cfg["map"]]))
    extract_media(pptx_path, med_dir, key_slides)

    # 3. Find image files for each section
    def find_img(si, prefer_tall=False):
        """Find best image for slide si"""
        candidates = []
        for f in sorted(os.listdir(med_dir)):
            if f.startswith(f"s{si:02d}_") and not f.startswith(f"s{si:02d}_g"):
                p = os.path.join(med_dir, f)
                try:
                    img = Image.open(p)
                    candidates.append((p, img.size))
                except Exception:
                    pass
        if not candidates:
            return None
        # Prefer largest image
        candidates.sort(key=lambda x: x[1][0]*x[1][1], reverse=True)
        return candidates[0][0]

    def find_boq_imgs(si):
        """BOQ slide has 2 images: tall one = form, wide one = price"""
        candidates = []
        for f in sorted(os.listdir(med_dir)):
            if f.startswith(f"s{si:02d}_"):
                p = os.path.join(med_dir, f)
                try:
                    img = Image.open(p)
                    candidates.append((p, img.size[0], img.size[1]))
                except Exception:
                    pass
        if len(candidates) < 2:
            return (candidates[0][0] if candidates else None, None)
        # Tall = boq form, wide = price table
        tall = sorted(candidates, key=lambda x: x[2]/max(x[1],1), reverse=True)
        wide = sorted(candidates, key=lambda x: x[1]/max(x[2],1), reverse=True)
        return tall[0][0], wide[0][0]

    # 4. Render map slide
    if map_override and os.path.exists(map_override):
        print("Using map_override image")
        map_jpg = map_override
    else:
        print("Rendering map slide...")
        map_jpg = render_map_slide(pptx_path, cfg["map"], work_dir)
        if not map_jpg:
            map_jpg = find_img(cfg["map"])

    # 5. Get title
    title1, title2 = get_title_from_pptx(pptx_path)
    print(f"Title: {title1}")
    print(f"Location: {title2}")

    # 6. Logo is already set at function entry

    boq_img, price_img = find_boq_imgs(cfg["boq"])
    surv_img  = find_img(cfg["surv"])
    des_img   = find_img(cfg["des"])
    cross_img = find_img(cfg["cross"])
    vol_img   = find_img(cfg["vol"])
    lett1_img = find_img(2)
    lett2_img = find_img(3)

    # 7. Build board
    board = Image.new("RGB", (W, H), NAVY)
    draw  = ImageDraw.Draw(board)

    # Header
    draw.rectangle([MG, MG, W-MG, MG+HDR], fill=NAVY)
    logo_sz = 340

    # Logo — มุมบนซ้าย
    if logo_path and os.path.exists(logo_path):
        lg = fit(logo_path, logo_sz, logo_sz, NAVY)
        board.paste(lg, (XL, MG + (HDR-logo_sz)//2))

    # ชื่อหน่วยงาน — มุมบนขวา
    agency_line1 = "นพค.43  สนภ.4"
    agency_line2 = "นทพ."
    fa  = fnt(90, bold=True)
    fa2 = fnt(80, bold=True)
    bb_a1 = draw.textbbox((0,0), agency_line1, font=fa)
    bb_a2 = draw.textbbox((0,0), agency_line2, font=fa2)
    right_x = W - MG - logo_sz
    draw.text((right_x - (bb_a1[2]-bb_a1[0]), MG + 60),  agency_line1, font=fa,  fill=GOLD)
    draw.text((right_x - (bb_a2[2]-bb_a2[0]), MG + 175), agency_line2, font=fa2, fill=WHITE)

    t1sz = 116 if len(title1) < 60 else 104
    f1 = fnt(t1sz, True); f2 = fnt(90)
    cx = W // 2
    bb1 = draw.textbbox((0,0), title1, font=f1)
    bb2 = draw.textbbox((0,0), title2, font=f2)
    draw.text((cx-(bb1[2]-bb1[0])//2, MG+45),  title1, font=f1, fill=GOLD)
    draw.text((cx-(bb2[2]-bb2[0])//2, MG+205), title2, font=f2, fill=WHITE)

    # Left column
    boq_h = int(CONH * 0.54)
    sec(draw, board, "แบบสรุปรายงานขุดลอกลำน้ำ", boq_img,   XL, CONY, CL, boq_h)
    sec(draw, board, "รายละเอียดราคากลาง",         price_img, XL, CONY+boq_h+GAP, CL, CONH-boq_h-GAP)

    # Middle column
    map_h = int(CONH * 0.50)
    sec(draw, board, "แผนที่และจุดดำเนินการ (มาตราส่วน 1:50,000)",
        map_jpg, XM, CONY, CM, map_h, lsz=60)
    sy = CONY + map_h + GAP
    sh = int(CONH * 0.265)
    sec(draw, board, "ตารางการสำรวจ",    surv_img, XM, sy,          CM, sh)
    sec(draw, board, "ตารางการออกแบบ",   des_img,  XM, sy+sh+GAP,   CM, CONH-map_h-GAP-sh-GAP)

    # Right column
    ch = int(CONH * 0.375)
    sec(draw, board, "รูปตัดตามขวางขุดลอกลำน้ำ", cross_img, XR, CONY, CR, ch)

    ly = CONY + ch + GAP
    lh = int(CONH * 0.37)
    LW2 = (CR - GAP) // 2
    draw.rectangle([XR, ly, XR+CR, ly+LH], fill=NAVY)
    lbl = "หนังสือรับรองการดำเนินงาน / ตรวจสอบความซ้ำซ้อน"
    fl = fnt(58, True)
    bb = draw.textbbox((0,0), lbl, font=fl)
    draw.text((XR+(CR-(bb[2]-bb[0]))//2, ly+(LH-(bb[3]-bb[1]))//2),
              lbl, font=fl, fill=GOLD)
    if lett1_img: board.paste(fit(lett1_img, LW2, lh-LH), (XR,          ly+LH))
    if lett2_img: board.paste(fit(lett2_img, LW2, lh-LH), (XR+LW2+GAP, ly+LH))

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
        fp = fnt(76, True)
        bbl = draw.textbbox((0,0), lbl, font=fp)
        draw.text((px+(PW-(bbl[2]-bbl[0]))//2, phy+(LH-(bbl[3]-bbl[1]))//2),
                  lbl, font=fp, fill=GOLD)
        if ph and os.path.exists(ph):
            board.paste(fit(ph, PW, PHH-LH), (px, phy+LH))

    board.save(output_path, "JPEG", quality=95, dpi=(150, 150))
    print(f"Saved: {output_path}")
    return output_path
