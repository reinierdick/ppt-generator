from flask import Flask, request, send_file
from pptx import Presentation
import tempfile
import os
import zipfile
import uuid

app = Flask(__name__)

UPLOAD_DIR = "uploads"

os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.route("/")
def home():
    return """
    <h1>Portfolio Upload</h1>
    <a href="/upload">Upload ZIP</a>
    """

@app.route("/upload")
def upload_page():
    return """
    <h2>Upload Portfolio ZIP</h2>

    <form action="/upload_zip" method="post" enctype="multipart/form-data">
        <input type="file" name="zipfile">
        <input type="submit" value="Upload">
    </form>
    """

@app.route("/upload_zip", methods=["POST"])
def upload_zip():

    if "zipfile" not in request.files:
        return {"error": "No file uploaded"}, 400

    file = request.files["zipfile"]

    portfolio_id = str(uuid.uuid4())

    portfolio_folder = os.path.join(
        UPLOAD_DIR,
        portfolio_id
    )

    os.makedirs(portfolio_folder)

    zip_path = os.path.join(
        portfolio_folder,
        "portfolio.zip"
    )

    file.save(zip_path)

    with zipfile.ZipFile(zip_path, "r") as zip_ref:
        zip_ref.extractall(portfolio_folder)

    image_count = 0

    for root, dirs, files in os.walk(portfolio_folder):
        for f in files:
            if f.lower().endswith(
                (".jpg", ".jpeg", ".png", ".webp")
            ):
                image_count += 1

    return {
        "success": True,
        "portfolio_id": portfolio_id,
        "image_count": image_count
    }

@app.route("/create_ppt", methods=["POST"])
def create_ppt():

    data = request.json

    prs = Presentation()

    for slide_data in data["slides"]:

        slide = prs.slides.add_slide(
            prs.slide_layouts[1]
        )

        slide.shapes.title.text = slide_data["title"]

        slide.placeholders[1].text = "\n".join(
            slide_data.get("bullets", [])
        )

    tmp = tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".pptx"
    )

    prs.save(tmp.name)

    filename = os.path.basename(tmp.name)

    base_url = request.host_url.rstrip("/")

    return {
        "success": True,
        "filename": filename,
        "download_url": f"{base_url}/download/{filename}"
    }

@app.route("/download/<filename>")
def download_file(filename):

    filepath = f"/tmp/{filename}"

    if not os.path.exists(filepath):
        return {"error": "file not found"}, 404

    return send_file(
        filepath,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
```
