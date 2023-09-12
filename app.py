from flask import Flask, render_template, request, redirect, url_for, send_file
import os
from werkzeug.utils import secure_filename
import pandas as pd
from converter_2 import main

app = Flask(__name__)

# Configure the upload folder
UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"xlsx"}
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# Define the output folder
OUTPUT_FOLDER = "output"  # You can change this to your desired folder name
app.config[
    "OUTPUT_FOLDER"
] = OUTPUT_FOLDER  # Set the output folder in Flask's configuration


# Function to check allowed file extensions
def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    success_message = request.args.get(
        "success_message", ""
    )  # Get the success message from the query parameter
    return render_template(
        "index.html", success_message=success_message
    )  # Pass the success message to the template


@app.route("/upload", methods=["POST"])
def upload_file():
    if "users_file" not in request.files or "system_program_file" not in request.files:
        return "Both files are required"

    users_file = request.files["users_file"]
    system_program_file = request.files["system_program_file"]
    time_ranges = request.form.get("time_ranges", "")

    """if users_file.filename == "" or system_program_file.filename == "":
        return "Both files must be selected"
    """

    if (users_file and allowed_file(users_file.filename)) and (
        system_program_file and allowed_file(system_program_file.filename)
    ):
        # Save the uploaded files
        users_filename = secure_filename(users_file.filename)
        users_file_path = os.path.join(app.config["UPLOAD_FOLDER"], users_filename)
        users_file.save(users_file_path)

        system_program_filename = secure_filename(system_program_file.filename)
        system_program_file_path = os.path.join(
            app.config["UPLOAD_FOLDER"], system_program_filename
        )
        system_program_file.save(system_program_file_path)

        # Update the filenames in your script
        excel_file_reference = users_file_path
        excel_file = system_program_file_path

        # Run your script here with the uploaded files
        output_file = main(
            excel_file, excel_file_reference, app.config["OUTPUT_FOLDER"], time_ranges
        )

        # Define the path to the output file
        # output_file_path = os.path.join(app.config['OUTPUT_FOLDER'], output_file)

        # Process the result and display it on a webpage
        # Example: You can read the generated output file and convert it to HTML or another suitable format

        # Send the output file as a response for download
        return send_file(output_file, as_attachment=True)

    else:
        return "Both files must be in .xlsx format"


if __name__ == "__main__":
    app.run(debug=True)
