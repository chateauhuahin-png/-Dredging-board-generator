"""
Flask Web App - ระบบสร้างบอร์ดชี้แจง
"""
import os, uuid, shutil
from flask import Flask, request, send_file, render_template, jsonify
from board_builder import build_board

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB per chunk (generous)

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload/start", methods=["POST"])
def upload_start():
    """Start a chunked upload session."""
    upload_id = str(uuid.uuid4())[:8]
    session_dir = os.path.join(UPLOAD_DIR, upload_id)
    os.makedirs(session_dir, exist_ok=True)
    return jsonify({"upload_id": upload_id})


@app.route("/upload/chunk", methods=["POST"])
def upload_chunk():
    """Receive one chunk and save it."""
    upload_id   = request.form.get("upload_id")
    chunk_index = int(request.form.get("chunk_index", 0))
    chunk_file  = request.files.get("chunk")

    if not upload_id or not chunk_file:
        return jsonify({"error": "ข้อมูลไม่ครบ"}), 400

    session_dir = os.path.join(UPLOAD_DIR, upload_id)
    if not os.path.isdir(session_dir):
        return jsonify({"error": "ไม่พบ upload session"}), 404

    chunk_path = os.path.join(session_dir, f"chunk_{chunk_index:05d}")
    chunk_file.save(chunk_path)
    return jsonify({"ok": True})


@app.route("/generate", methods=["POST"])
def generate():
    """Assemble chunks and build board."""
    data = request.get_json()
    upload_id = data.get("upload_id") if data else None

    if not upload_id:
        return jsonify({"error": "ไม่พบ upload_id"}), 400

    session_dir = os.path.join(UPLOAD_DIR, upload_id)
    if not os.path.isdir(session_dir):
        return jsonify({"error": "ไม่พบไฟล์อัปโหลด"}), 404

    pptx_path   = os.path.join(session_dir, "input.pptx")
    output_path = os.path.join(OUTPUT_DIR, f"board_{upload_id}.jpg")

    try:
        # Assemble chunks in order
        chunks = sorted(
            f for f in os.listdir(session_dir) if f.startswith("chunk_")
        )
        if not chunks:
            return jsonify({"error": "ไม่พบ chunk ไฟล์"}), 400

        with open(pptx_path, "wb") as out:
            for chunk_name in chunks:
                with open(os.path.join(session_dir, chunk_name), "rb") as cf:
                    out.write(cf.read())

        build_board(
            pptx_path=pptx_path,
            work_dir=session_dir,
            output_path=output_path,
        )

        return jsonify({"upload_id": upload_id, "url": f"/download/{upload_id}"})

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        try:
            shutil.rmtree(session_dir, ignore_errors=True)
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
