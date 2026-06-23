"""
Flask Web App - ระบบสร้างบอร์ดชี้แจง
"""
import os, uuid, shutil
from flask import Flask, request, send_file, render_template, jsonify
from board_builder import build_board

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


def save_upload(file, dest_dir, filename):
    if not file or file.filename == "":
        return None
    path = os.path.join(dest_dir, filename)
    file.save(path)
    return path


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    job_id = str(uuid.uuid4())[:8]
    work_dir = os.path.join(UPLOAD_DIR, job_id)
    os.makedirs(work_dir, exist_ok=True)

    try:
        pptx_file   = request.files.get("pptx")
        photo_before = request.files.get("photo_before")
        photo_during = request.files.get("photo_during")
        photo_after  = request.files.get("photo_after")
        logo_file    = request.files.get("logo")

        if not pptx_file:
            return jsonify({"error": "กรุณาอัปโหลดไฟล์ PPTX"}), 400

        pptx_path    = save_upload(pptx_file,    work_dir, "input.pptx")
        before_path  = save_upload(photo_before,  work_dir, "before.jpg")
        during_path  = save_upload(photo_during,  work_dir, "during.jpg")
        after_path   = save_upload(photo_after,   work_dir, "after.jpg")
        logo_path    = save_upload(logo_file,     work_dir, "logo.png") if logo_file and logo_file.filename else None

        output_path = os.path.join(OUTPUT_DIR, f"board_{job_id}.jpg")

        build_board(
            pptx_path=pptx_path,
            photo_before=before_path,
            photo_during=during_path,
            photo_after=after_path,
            work_dir=work_dir,
            output_path=output_path,
            logo_path=logo_path,
        )

        return jsonify({"job_id": job_id, "url": f"/download/{job_id}"})

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        # Clean up work files (keep output)
        try:
            shutil.rmtree(work_dir, ignore_errors=True)
        except Exception:
            pass


@app.route("/download/<job_id>")
def download(job_id):
    path = os.path.join(OUTPUT_DIR, f"board_{job_id}.jpg")
    if not os.path.exists(path):
        return "ไม่พบไฟล์", 404
    return send_file(path, as_attachment=True,
                     download_name="บอร์ดชี้แจง.jpg",
                     mimetype="image/jpeg")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
