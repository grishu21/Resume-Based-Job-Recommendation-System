from werkzeug.utils import secure_filename
import os

@app.route("/upload_proofs", methods=["POST"])
def upload_proofs():
    tech_proof = request.files.get("tech_proof")
    nontech_proof = request.files.get("nontech_proof")

    tech_text, nontech_text = "", ""

    # Handle Technical Proof
    if tech_proof and tech_proof.filename != "":
        filename = secure_filename(tech_proof.filename)
        tech_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        
        # Save file first
        tech_proof.save(tech_path)

        # Now extract text
        tech_text = extract_text_from_file(tech_path)

        # Optional: delete file after extraction
        if os.path.exists(tech_path):
            os.remove(tech_path)

    # Handle Non-Technical Proof
    if nontech_proof and nontech_proof.filename != "":
        filename = secure_filename(nontech_proof.filename)
        nontech_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)

        # Save file first
        nontech_proof.save(nontech_path)

        # Now extract text
        nontech_text = extract_text_from_file(nontech_path)

        # Optional: delete file after extraction
        if os.path.exists(nontech_path):
            os.remove(nontech_path)

    # Merge extracted text for skill extraction
    combined_text = tech_text + "\n" + nontech_text
    skills = extract_skills(combined_text)

    return render_template("result.html", skills=skills)
