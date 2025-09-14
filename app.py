import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import zipfile
from flask import Flask, render_template, request, send_file, jsonify

app = Flask(__name__)

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "certificate.png")
FONT_PATH_BOLD = os.path.join(BASE_DIR, "fonts", "Rillosta.ttf")
FONT_PATH_REGULAR = os.path.join(BASE_DIR, "fonts", "OpenSans-Regular.ttf")
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "generated_output")

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    excel_file = request.files.get("excel")
    template_file = request.files.get("template")

    if not excel_file or not template_file:
        return jsonify({"error": "Both Excel and Template are required!"}), 400

    # Save uploaded files
    excel_path = os.path.join(INPUT_DIR, "data.xlsx")
    excel_file.save(excel_path)

    template_path = TEMPLATE_PATH
    template_file.save(template_path)

    # Run generation
    try:
        zip_path = generate_certificates(excel_path, template_path)
        return jsonify({"message": "Certificates generated successfully!", "zip": "/download"})
    except Exception as e:
        return jsonify({"error": f"Error generating certificates: {str(e)}"}), 500


@app.route("/download")
def download():
    zip_path = os.path.join(OUTPUT_DIR, "certificates.zip")
    if os.path.exists(zip_path):
        return send_file(zip_path, as_attachment=True)
    return jsonify({"error": "No certificates generated yet!"}), 404


def ordinal(n):
    try:
        n = int(n)
        if 11 <= (n % 100) <= 13:
            suffix = 'th'
        else:
            suffix = ['th', 'st', 'nd', 'rd', 'th'][min(n % 10, 4)]
        return str(n) + suffix
    except ValueError:
        return str(n)


def generate_certificates(excel_path, template_path):
    # Read Excel file
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        raise ValueError(f"Error reading Excel file: {str(e)}")

    # Drop rows where Name is NaN or empty
    df = df.dropna(subset=["Name"])
    df = df[df["Name"].str.strip() != ""]

    # Verify required columns
    required_columns = ["Name", "Course", "Position", "Event"]
    if not all(col in df.columns for col in required_columns):
        missing_cols = [col for col in required_columns if col not in df.columns]
        raise ValueError(f"Missing required columns in Excel: {missing_cols}")

    # Load template
    try:
        template = Image.open(template_path)
        width, height = template.size
        print(f"Template image loaded: {width}x{height}")
    except Exception as e:
        raise ValueError(f"Error loading template image: {str(e)}")

    # Load fonts with fallback
    try:
        name_font = ImageFont.truetype(FONT_PATH_BOLD, 130)
        details_font = ImageFont.truetype(FONT_PATH_REGULAR, 50)
    except Exception as e:
        print(f"Font loading error: {str(e)}. Using default font.")
        name_font = ImageFont.load_default()
        details_font = ImageFont.load_default()

    # Clean output directory
    for f in os.listdir(OUTPUT_DIR):
        os.remove(os.path.join(OUTPUT_DIR, f))

    # Generate certificates
    for index, row in df.iterrows():
        name = str(row["Name"]) if pd.notna(row["Name"]) else "Unknown"
        course = str(row["Course"]) if pd.notna(row["Course"]) else "Unknown"
        position_raw = row["Position"] if pd.notna(row["Position"]) else "Unknown"
        event = str(row["Event"]) if pd.notna(row["Event"]) else "Unknown"

        # Handle position
        position_str = str(position_raw).strip()
        if position_str.isdigit():
            position = ordinal(position_raw)
        else:
            position = position_str

        print(f"Generating certificate {index + 1}: Name={name}, Course={course}, Position={position}, Event={event}")

        cert = template.copy()
        draw = ImageDraw.Draw(cert)

        # Fixed coordinates (adjust these as needed based on your template)
        name_coord = (750, 570)  # Example: Centered for name (big blank?)
        course_coord = (1560, 580)  # Example: For course, perhaps after "of"
        position_coord = (980, 670)  # Example: After "secured"
        event_coord = (1590, 680)  # Example: After "Position in"

        # Log coordinates
        print(f"Text coordinates: Name={name_coord}, Course={course_coord}, Position={position_coord}, Event={event_coord}")

        # Draw text with appropriate anchors
        draw.text(name_coord, name, font=name_font, fill="black", anchor="mm")  # Centered
        draw.text(course_coord, course, font=details_font, fill="black", anchor="mm")  # Centered
        draw.text(position_coord, position, font=details_font, fill="black", anchor="ms")  # Left-aligned
        draw.text(event_coord, event, font=details_font, fill="black", anchor="ms")  # Left-aligned

        # Save certificate
        safe_name = name.replace(' ', '_').replace('.', '')
        output_path = os.path.join(OUTPUT_DIR, f"{safe_name}_certificate.png")
        cert.save(output_path)
        print(f"Saved certificate: {output_path}")

    # Create zip
    zip_path = os.path.join(OUTPUT_DIR, "certificates.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in os.listdir(OUTPUT_DIR):
            if file.endswith(".png"):
                zipf.write(os.path.join(OUTPUT_DIR, file), arcname=file)
        print(f"Created zip file: {zip_path}")

    return zip_path


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)