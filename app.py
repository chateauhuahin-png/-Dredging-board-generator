"""
Flask Web App - ระบบสร้างบอร์ดชี้แจง
"""
import os, uuid, shutil, threading
from flask import Flask, request, send_file, render_template, jsonify
from board_builder import build_board

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB per chunk

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# In-memory job status store
jobs = {}  # job_id -> {"status": "processing"|"done"|"error", "error": str}


@app.route("/")
def index():
    return render_template("index.html")


# ── Chunked upload ────────────────────────────────────────────────────────────

@app.route("/upload/start", methods=["POST"])
def upload_start():
    upload_id = str(uuid.uuid4())[:8]
    session_dir = os.path.join(UPLOAD_DIR, upload_id)
    os.makedirs(session_dir, exist_ok=True)
    return jsonify({"upload_id": upload_id})


@app.route("/upload/chunk", methods=["POST"])
def upload_chunk():
    upload_id   = request.form.get("upload_id")
    chunk_index = int(request.form.get("chunk_index", 0))
    chunk_file  = request.files.get("chunk")

    if not upload_id or not chunk_file:
        return jsonify({"error": "ข้อมูลไม่ครบ"}), 400

    session_dir = os.path.join(UPLOAD_DIR, upload_id)
    if not os.path.isdir(session_dir):
        return jsonify({"error": "ไม่พบ upload session"}), 404

    chunk_file.save(os.path.join(session_dir, f"chunk_{chunk_index:05d}"))
    return jsonify({"ok": True})


# ── Async generate ────────────────────────────────────────────────────────────

def _run_build(upload_id):
    session_dir = os.path.join(UPLOAD_DIR, upload_id)
    pptx_path   = os.path.join(session_dir, "input.pptx")
    output_path = os.path.join(OUTPUT_DIR, f"board_{upload_id}.jpg")
    try:
        # Assemble chunks
        chunks = sorted(f for f in os.listdir(session_dir) if f.startswith("chunk_"))
        with open(pptx_path, "wb") as out:
            for c in chunks:
                with open(os.path.join(session_dir, c), "rb") as cf:
                    out.write(cf.read())

        build_board(pptx_path=pptx_path, work_dir=session_dir, output_path=output_path)
        jobs[upload_id] = {"status": "done"}
    except Exception as e:
        jobs[upload_id] = {"status": "error", "error": str(e)}
    finally:
        shutil.rmtree(session_dir, ignore_errors=True)


@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json(force=True, silent=True) or {}
        upload_id = data.get("upload_id")

        if not upload_id:
            return jsonify({"error": "ไม่พบ upload_id"}), 400

        session_dir = os.path.join(UPLOAD_DIR, upload_id)
        if not os.path.isdir(session_dir):
            return jsonify({"error": f"ไม่พบ session dir: {session_dir}"}), 404

        jobs[upload_id] = {"status": "processing"}
        t = threading.Thread(target=_run_build, args=(upload_id,), daemon=True)
        t.start()

        return jsonify({"job_id": upload_id})
    except Exception as e:
        return jsonify({"error": f"generate error: {str(e)}"}), 500


@app.route("/status/<job_id>")
def status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"status": "not_found"}), 404
    if job["status"] == "done":
        return jsonify({"status": "done", "url": f"/download/{job_id}"})
    if job["status"] == "error":
        return jsonify({"status": "error", "error": job.get("error", "ไม่ทราบสาเหตุ")})
    return jsonify({"status": "processing"})


# ── Download ──────────────────────────────────────────────────────────────────

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
