from flask import Flask, request, send_file
from pptx import Presentation
from pptx.util import Inches
import tempfile
import os
import zipfile
import uuid
import json

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


@app.route("/portfolio/<portfolio_id>")
def portfolio_info(portfolio_id):

    portfolio_folder = os.path.join(
        UPLOAD_DIR,
        portfolio_id
    )

    if not os.path.exists(portfolio_folder):
        return {"error": "portfolio not found"}, 404

    files = []

    for root, dirs, filenames in os.walk(portfolio_folder):
        for f in filenames:
            files.append(
                os.path.relpath(
                    os.path.join(root, f),
                    portfolio_folder
                )
            )

    return {
        "portfolio_id": portfolio_id,
        "files": files
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

    if base_url.startswith("http://"):
        base_url = base_url.replace(
            "http://",
            "https://",
            1
        )

    return {
        "success": True,
        "filename": filename,
        "download_url": f"{base_url}/download/{filename}"
    }


@app.route("/create_portfolio_ppt/<portfolio_id>")
def create_portfolio_ppt(portfolio_id):

    portfolio_folder = os.path.join(
        UPLOAD_DIR,
        portfolio_id
    )

    if not os.path.exists(portfolio_folder):
        return {"error": "portfolio not found"}, 404

    prs = Presentation()

    image_extensions = (
        ".jpg",
        ".jpeg",
        ".png",
        ".webp"
    )

    images = []

    for file in os.listdir(portfolio_folder):

        if file.lower().endswith(
            image_extensions
        ):
            images.append(
                os.path.join(
                    portfolio_folder,
                    file
                )
            )

    images.sort()

    title_slide = prs.slides.add_slide(
        prs.slide_layouts[0]
    )

    title_slide.shapes.title.text = (
        f"Portfolio {portfolio_id}"
    )

    title_slide.placeholders[1].text = (
        f"{len(images)} foto's"
    )

    for image_path in images:

        slide = prs.slides.add_slide(
            prs.slide_layouts[5]
        )

        try:
            slide.shapes.title.text = (
                os.path.basename(image_path)
            )
        except:
            pass

        slide.shapes.add_picture(
            image_path,
            Inches(0.5),
            Inches(1),
            width=Inches(8)
        )

    tmp = tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".pptx"
    )

    prs.save(tmp.name)

    filename = os.path.basename(
        tmp.name
    )

    base_url = request.host_url.rstrip("/")

    return {
        "success": True,
        "portfolio_id": portfolio_id,
        "image_count": len(images),
        "download_url":
            f"{base_url}/download/{filename}"
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

@app.route("/portfolio/<portfolio_id>/images")
def portfolio_images(portfolio_id):

    portfolio_folder = os.path.join(
        UPLOAD_DIR,
        portfolio_id
    )

    if not os.path.exists(portfolio_folder):
        return {"error": "portfolio not found"}, 404

    image_extensions = (
        ".jpg",
        ".jpeg",
        ".png",
        ".webp"
    )

    base_url = request.host_url.rstrip("/")

    images = []

    for file in os.listdir(portfolio_folder):

        if file.lower().endswith(image_extensions):

            images.append({
                "filename": file,
                "url": f"{base_url}/image/{portfolio_id}/{file}"
            })

    return {
        "portfolio_id": portfolio_id,
        "image_count": len(images),
        "images": images
    }


@app.route("/image/<portfolio_id>/<filename>")
def serve_image(portfolio_id, filename):

    filepath = os.path.join(
        UPLOAD_DIR,
        portfolio_id,
        filename
    )

    if not os.path.exists(filepath):
        return {"error": "image not found"}, 404

    return send_file(filepath)
@app.route("/portfolio/<portfolio_id>/manifest")
def get_manifest(portfolio_id):

    manifest_path = os.path.join(
        UPLOAD_DIR,
        portfolio_id,
        "manifest.json"
    )

    if not os.path.exists(manifest_path):
        return {
            "portfolio_id": portfolio_id,
            "analyses": []
        }

    with open(manifest_path, "r") as f:
        return json.load(f)


@app.route(
    "/portfolio/<portfolio_id>/analysis",
    methods=["POST"]
)
def save_analysis(portfolio_id):

    portfolio_folder = os.path.join(
        UPLOAD_DIR,
        portfolio_id
    )

    if not os.path.exists(portfolio_folder):
        return {"error": "portfolio not found"}, 404

    data = request.json

    manifest_path = os.path.join(
        portfolio_folder,
        "manifest.json"
    )

    if os.path.exists(manifest_path):

        with open(manifest_path, "r") as f:
            manifest = json.load(f)

    else:

        manifest = {
            "portfolio_id": portfolio_id,
            "analyses": []
        }

    manifest["analyses"].append(data)

    with open(manifest_path, "w") as f:
        json.dump(
            manifest,
            f,
            indent=2
        )

    return {
        "success": True,
        "analysis_count":
            len(manifest["analyses"])
    }
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
