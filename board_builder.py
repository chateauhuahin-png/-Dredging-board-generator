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


def render_map_slide(pptx_path, slide_idx, work_dir):
    """Render slide to JPG using PIL — exact EMU coordinates, pictures + text + lines"""
    import io as _io
    from lxml import etree

    NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

    prs   = Presentation(pptx_path)
    slide = prs.slides[slide_idx - 1]
    sw    = prs.slide_width
    sh    = prs.slide_height

    OUT_W = 1800
    scale = OUT_W / sw
    OUT_H = int(sh * scale)

    canvas = Image.new("RGB", (OUT_W, OUT_H), (255, 255, 255))
    draw   = ImageDraw.Draw(canvas)

    def px(emu):
        return int((emu or 0) * scale)

    def get_rgb(color_elem):
        """Try to extract (r,g,b) from an lxml color element"""
        try:
            val = color_elem.get("val") or color_elem.get("lastClr")
            if val:
                v = int(val, 16)
                return ((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF)
        except Exception:
            pass
        return None

    def line_color_width(shape):
        """Return (color, width_px) for a shape's outline"""
        color = (0, 0, 0)
        lw = max(1, px(12700))  # default ~0.5pt
        try:
            ln = shape._element.spPr.find(f"{{{NS}}}ln")
            if ln is None:
                ln = shape._element.find(f".//{{{NS}}}ln")
            if ln is not None:
                w_emu = ln.get("w")
                if w_emu:
                    lw = max(1, px(int(w_emu)))
                solid = ln.find(f"{{{NS}}}solidFill")
                if solid is not None:
                    for tag in ("srgbClr", "sysClr"):
                        el = solid.find(f"{{{NS}}}{tag}")
                        if el is not None:
                            c = get_rgb(el)
                            if c:
                                color = c
                                break
        except Exception:
            pass
        return color, lw

    def is_line_shape(shape):
        """True if shape is drawn as a line/arrow (not a filled block)"""
        try:
            spPr = shape._element.spPr
            pg = spPr.find(f"{{{NS}}}prstGeom")
            if pg is not None:
                prst = pg.get("prst", "")
                LINE_PRSTS = {"line", "lineInv", "straightConnector1",
                              "bentConnector2", "bentConnector3",
                              "curvedConnector2", "curvedConnector3",
                              "leftRightArrow", "upDownArrow",
                              "rightArrow", "leftArrow", "upArrow", "downArrow"}
                if prst in LINE_PRSTS or "line" in prst.lower() or "Arrow" in prst:
                    return True
        except Exception:
            pass
        return False

    def draw_line_shape(shape, x, y, w, h):
        color, lw = line_color_width(shape)
        try:
            xfrm = shape._element.spPr.xfrm
            flipH = xfrm is not None and xfrm.get("flipH") == "1"
            flipV = xfrm is not None and xfrm.get("flipV") == "1"
        except Exception:
            flipH = flipV = False
        x1, y1 = (x + w if flipH else x), (y + h if flipV else y)
        x2, y2 = (x if flipH else x + w), (y if flipV else y + h)
        draw.line([(x1, y1), (x2, y2)], fill=color, width=lw)

    def draw_text_frame(shape, x, y, w, h):
        try:
            tf = shape.text_frame
        except Exception:
            return
        ty = y
        for para in tf.paragraphs:
            line_texts = []
            sz_px = 16
            color  = (0, 0, 0)
            bold   = False
            for run in para.runs:
                if not run.text:
                    continue
                try:
                    if run.font.size:
                        sz_px = max(8, int(run.font.size.pt * scale * 96 / 72))
                except Exception:
                    pass
                try:
                    bold = bool(run.font.bold)
                except Exception:
                    pass
                try:
                    rgb = run.font.color.rgb
                    color = (rgb.red, rgb.green, rgb.blue)
                except Exception:
                    pass
                line_texts.append(run.text)
            text = "".join(line_texts).strip()
            if text:
                f = fnt(sz_px, bold)
                draw.text((x + 4, ty), text, font=f, fill=color)
                ty += int(sz_px * 1.3)
            else:
                ty += int(sz_px * 0.6)

    def render_shapes(shapes, dx=0, dy=0):
        for shape in shapes:
            try:
                x = dx + px(shape.left)
                y = dy + px(shape.top)
                w = px(shape.width)
                h = px(shape.height)
            except Exception:
                continue

            stype = shape.shape_type

            if stype == 13:                      # Picture
                try:
                    img = Image.open(_io.BytesIO(shape.image.blob)).convert("RGBA")
                    if w > 0 and h > 0:
                        img = img.resize((w, h), Image.LANCZOS)
                        bg  = Image.new("RGB", (w, h), (255, 255, 255))
                        bg.paste(img, mask=img.split()[3])
                        canvas.paste(bg, (x, y))
                        del img, bg
                except Exception as e:
                    print(f"  pic error: {e}")

            elif stype == 6:                     # Group — recurse
                render_shapes(shape.shapes, x, y)

            elif stype == 9:                     # Connector / line
                color, lw = line_color_width(shape)
                draw.line([(x, y), (x + w, y + h)], fill=color, width=lw)

            else:                                # Auto shape / text box
                if is_line_shape(shape):
                    draw_line_shape(shape, x, y, w, h)
                if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                    draw_text_frame(shape, x, y, w, h)

    render_shapes(slide.shapes)

    jpg = os.path.join(work_dir, "map_slide.jpg")
    canvas.save(jpg, "JPEG", quality=92)
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
                work_dir, output_path, logo_path=None):
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

    # 4. Render map slide via LibreOffice
    print("Rendering map slide...")
    map_jpg = render_map_slide(pptx_path, cfg["map"], work_dir)
    if not map_jpg:
        map_jpg = find_img(cfg["map"])

    # 5. Get title
    title1, title2 = get_title_from_pptx(pptx_path)
    print(f"Title: {title1}")
    print(f"Location: {title2}")

    # 6. Get logo
    if not logo_path:
        logo_path = os.path.join(BASE_DIR, "fonts", "logo.png")  # optional

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
    if logo_path and os.path.exists(logo_path):
        lg = fit(logo_path, logo_sz, logo_sz, NAVY)
        board.paste(lg, (XL, MG + (HDR-logo_sz)//2))
        board.paste(lg, (W-MG-logo_sz, MG + (HDR-logo_sz)//2))

    t1sz = 116 if len(title1) < 60 else 104
    f1 = fnt(t1sz, True); f2 = fnt(90)
    cx = W // 2
    bb1 = draw.textbbox((0,0), title1, font=f1)
    bb2 = draw.textbbox((0,0), title2, font=f2)
    draw.text((cx-(bb1[2]-bb1[0])//2, MG+45),  title1, font=f1, fill=GOLD)
    draw.text((cx-(bb2[2]-bb2[0])//2, MG+205), title2, font=f2, fill=WHITE)
    draw.text((W-MG-400, MG+HDR-64), "นทพ.", font=fnt(56), fill=WHITE)

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
