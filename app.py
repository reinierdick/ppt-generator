from flask import Flask, request, jsonify
from pptx import Presentation
import uuid

app = Flask(__name__)

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

    filename = f"{uuid.uuid4()}.pptx"

    prs.save(filename)

    return jsonify({
        "message": "ppt created",
        "filename": filename
    })

@app.route("/")
def home():
    return "PPT Generator Running"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
