"""
Microsoft Graph API — แปลง PPTX slide เป็น PNG ด้วย PowerPoint renderer จริง
ตั้งค่า env vars: AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_REFRESH_TOKEN
"""
import os, requests, subprocess, glob

CLIENT_ID     = os.environ.get("AZURE_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")
REFRESH_TOKEN = os.environ.get("AZURE_REFRESH_TOKEN", "")


def is_configured():
    return bool(CLIENT_ID and CLIENT_SECRET and REFRESH_TOKEN)


def get_access_token():
    r = requests.post(
        "https://login.microsoftonline.com/common/oauth2/v2.0/token",
        data={
            "client_id":     CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "refresh_token": REFRESH_TOKEN,
            "grant_type":    "refresh_token",
            "scope":         "https://graph.microsoft.com/Files.ReadWrite offline_access",
        }, timeout=20
    )
    data = r.json()
    if "access_token" not in data:
        raise Exception(f"Token error: {data.get('error_description', data)}")
    return data["access_token"]


def slide_to_png(pptx_path, slide_idx, work_dir):
    """
    อัปโหลด PPTX → OneDrive → ดาวน์โหลดเป็น PDF (PowerPoint render) → แยก slide เป็น PNG
    คืนค่า path ของ PNG หรือ None ถ้าล้มเหลว
    """
    if not is_configured():
        return None

    try:
        token   = get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        fname   = f"tmp_board_{os.getpid()}.pptx"

        # 1. อัปโหลด PPTX ไป OneDrive
        with open(pptx_path, "rb") as f:
            resp = requests.put(
                f"https://graph.microsoft.com/v1.0/me/drive/root:/{fname}:/content",
                headers={
                    **headers,
                    "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                },
                data=f, timeout=120
            )
        item    = resp.json()
        item_id = item.get("id")
        if not item_id:
            print(f"  graph upload failed: {item}")
            return None

        try:
            # 2. ดาวน์โหลดเป็น PDF (Microsoft ใช้ PowerPoint render จริง)
            pdf_resp = requests.get(
                f"https://graph.microsoft.com/v1.0/me/drive/items/{item_id}/content?format=pdf",
                headers=headers, allow_redirects=True, timeout=120
            )
            if pdf_resp.status_code != 200:
                print(f"  graph pdf error: {pdf_resp.status_code}")
                return None

            pdf_path = os.path.join(work_dir, "map_graph.pdf")
            with open(pdf_path, "wb") as f:
                f.write(pdf_resp.content)

            # 3. แยก slide ที่ต้องการเป็น PNG ด้วย pdftoppm
            out_prefix = os.path.join(work_dir, "map_graph")
            subprocess.run(
                ["pdftoppm", "-r", "200",
                 "-f", str(slide_idx), "-l", str(slide_idx),
                 "-png", pdf_path, out_prefix],
                capture_output=True, timeout=30
            )

            pngs = sorted(glob.glob(f"{out_prefix}*.png"))
            return pngs[0] if pngs else None

        finally:
            # 4. ลบไฟล์ออกจาก OneDrive
            requests.delete(
                f"https://graph.microsoft.com/v1.0/me/drive/items/{item_id}",
                headers=headers, timeout=15
            )

    except Exception as e:
        print(f"  graph convert error: {e}")
        return None
