import os
import json
import time
from io import BytesIO

from flask import Flask, render_template, request, send_file, flash, redirect
from werkzeug.utils import secure_filename
from google import genai
from openpyxl import Workbook
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = "change-this-secret"
app.config["UPLOAD_FOLDER"] = "uploads"
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

# ==============================
# Multiple Gemini API Keys
# ==============================

# Try multiple keys first
API_KEYS = os.getenv("GEMINI_API_KEYS")

# If not found, use single key
if not API_KEYS:
    single_key = os.getenv("GEMINI_API_KEY")

    if not single_key:
        raise RuntimeError(
            "No Gemini API key found in .env"
        )

    API_KEYS = [single_key]

else:
    API_KEYS = [k.strip() for k in API_KEYS.split(",") if k.strip()]

current_key = 0


def get_client():
    global current_key
    return genai.Client(api_key=API_KEYS[current_key])


def generate_with_fallback(prompt, uploaded_file):
    """
    Automatically switches to the next API key if quota is exceeded.
    """

    global current_key

    while current_key < len(API_KEYS):

        try:
            client = get_client()

            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=[prompt, uploaded_file],
            )

            return response

        except Exception as e:

            error = str(e).lower()
            print(error)

            if (
                "429" in error
                or "quota" in error
                or "resource_exhausted" in error
                or "rate limit" in error
            ):

                print(f"\nAPI Key {current_key+1} exhausted.")
                current_key += 1

                if current_key < len(API_KEYS):
                    print(f"Switching to API Key {current_key+1}...\n")
                    time.sleep(1)
                    continue

                raise Exception("All API Keys are exhausted.")

            raise


MODEL_NAME = "gemini-2.5-flash"

FIELDS = [
    "Name",
    "Class",
    "Admission Number",
    "phone Number",
    "Date of Birth",
    "Blood Group",
    "Birth Place",
    "Parent/Guardian Name",
    "Address",
]

EXTRACTION_PROMPT = f"""
You are reading a scanned student information card PDF.

The PDF may contain multiple student cards.

Extract every student's details.

Return ONLY valid JSON.

JSON Keys:
{json.dumps(FIELDS)}

Rules:
- One object per student.
- Missing field = "".
- Address in one line.
- No markdown.
- No explanation.
"""

ALLOWED_EXTENSIONS = {"pdf"}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_records_from_pdf(filepath):

    client = get_client()

    uploaded_file = client.files.upload(file=filepath)

    response = generate_with_fallback(
        EXTRACTION_PROMPT,
        uploaded_file
    )

    text = response.text.strip()

    if text.startswith("```"):
        text = text.replace("```json", "")
        text = text.replace("```", "").strip()

    try:
        data = json.loads(text)

        if isinstance(data, dict):
            data = [data]

    except Exception:
        data = [{
            **{field: "" for field in FIELDS},
            "_raw_response": text
        }]

    return data


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():

    if "file" not in request.files:
        flash("No file selected")
        return redirect("/")

    file = request.files["file"]

    if file.filename == "":
        flash("No file selected")
        return redirect("/")

    if not allowed_file(file.filename):
        flash("Upload PDF only")
        return redirect("/")

    filename = file.filename
    filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)

    file.save(filepath)

    try:
        records = extract_records_from_pdf(filepath)

    finally:
        if os.path.exists(filepath):
            os.remove(filepath)

    wb = Workbook()
    ws = wb.active
    ws.title = "Student Records"

    ws.append(FIELDS)

    for record in records:
        ws.append([record.get(field, "") for field in FIELDS])

    for column in ws.columns:
        length = max(len(str(cell.value or "")) for cell in column)
        ws.column_dimensions[column[0].column_letter].width = length + 5

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="student_records.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
