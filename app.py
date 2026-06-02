from flask import Flask, render_template, request, redirect, session, send_file, after_this_request
from docxtpl import DocxTemplate, InlineImage
from datetime import timedelta
from werkzeug.utils import secure_filename #for making the file name secure like my file.csv -> my_file.csv automatically
from PIL import Image
from docx.shared import Mm
import tempfile
import pandas as pd
import os
import json
import shutil

app = Flask(__name__)
app.secret_key = os.getenv("SECRETKEY")




@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    UPLOAD_FOLDER = tempfile.mkdtemp()
    session["UPLOAD_FOLDER"] = UPLOAD_FOLDER

    # =========================
    # Event Details
    # =========================
    code = request.form.get("code")
    date = request.form.get("date")
    duration = request.form.get("duration")
    venue = request.form.get("venue")
    register1 = request.form.get("register1")
    register2 = request.form.get("register2")
    register3 = request.form.get("register3")

    # =========================
    # CSV File
    # =========================
    csv_file = request.files.get("csv_file")

    if not csv_file:
        return "❌ No CSV file uploaded."

    # Save uploaded file
    filename = secure_filename(csv_file.filename)       #to making sure the file name has safe saving name 
    csv_path = os.path.join(UPLOAD_FOLDER, filename)        # it join the temp folder path with the secured file name
    csv_file.save(csv_path)     #save the file into the temp folder path

    # =========================
    # Read CSV
    # =========================
    try:
        df = pd.read_csv(csv_path)
    except Exception as e:
        return f"❌ Error reading CSV: {e}"


    # =========================
    # Validate Columns
    # =========================
    required_columns = ["Name", "RollNo"]

    for col in required_columns:
        if col not in df.columns:
            return f"❌ Missing required column: {col}"


    # =========================
    # Convert to List
    # =========================
    participants = df.to_dict(orient="records")

    # =========================
    # Store in Session
    # =========================
    session["event_details"] = {
        "code": code,
        "date": date,
        "duration": duration,
        "venue": venue,
        "register1": register1,
        "register2": register2,
        "register3": register3,

    }

    session["participants"] = participants

    # =========================
    # Debug Print
    # =========================
    # print("\n=== EVENT DETAILS ===")
    # print(session["event_details"])

    # print("\n=== PARTICIPANTS ===")
    # for p in participants:
    #     print(p)

    return redirect("/teams")

@app.route("/teams")
def teams():
    participants = session.get("participants", [])

    return render_template(
        "teams.html",
        participants=participants)


@app.route("/save_teams", methods=["POST"])
def save_teams():
    doc = DocxTemplate("template.docx")
    
    UPLOAD_FOLDER = session.get("UPLOAD_FOLDER")     #open the template format file

    teams_json = request.form.get("teams_data")     #getting the team data

    if not teams_json:
        return "❌ No team data received."

    teams = json.loads(teams_json)

    # ==========================
    # Final Context
    # ==========================

    context = session.get("event_details", {}).copy() #getting the event static data like venue and date

    context["teams"] = []   #creating the empty list of the team data then merge the team data with image in inline format

    # print("\n✅ Processing teams with images...")

    for i, team in enumerate(teams):

        # =========================
        # Get uploaded image
        # =========================
        file = request.files.get(f"mvp_{i}")

        if file and file.filename != "":

            # Get original extension
            ext = os.path.splitext(file.filename)[1].lower()

            filename = f"mvp_team_{i}{ext}"
            path = os.path.join(UPLOAD_FOLDER, filename)    #to merge the path of the temp folder with the image name

            # modify and correcting the format of the inserted image 
            img = Image.open(file.stream)
            img.convert("RGB").save(path, optimize=True)


            #convert to DOCX inline image
            image_obj = InlineImage(doc, path, width=Mm(102))
        else:
            image_obj = ""  




            

        # =========================
        # Build team context
        # =========================
        context["teams"].append({
            "team_no": team["team_no"],
            "team_name": team["team_name"],
            "ps_no": team["ps_no"],
            "length": team["length"],
            "members": team["members"],
            "mvp": image_obj
        })

    
    # for debbuging
    # print("\n=== FINAL CONTEXT ===")
    # print(json.dumps(context, indent=4))


    # 1. Get the safe data from the user


    print("\n✅ Data collection complete. Generating Word Document...")

    # 2. Pass it directly into your template
    doc.render(context)

    output_path = os.path.join(UPLOAD_FOLDER,"Hackathon_Report.docx") # adding the path of the temp folder with the report stored temporarly

    doc.save(output_path)
    # doc.save("REPORT.docx")
 
        # =========================
    # CLEANUP AFTER RESPONSE
    # =========================
    @after_this_request
    def cleanup(response):
        try:
            shutil.rmtree(UPLOAD_FOLDER, ignore_errors=True)
        except Exception as e:
            print("Cleanup error:", e)
        return response

    return send_file(
        output_path,    
        as_attachment=True,
        download_name="Report.docx"
    )


@app.route("/generate", methods=["POST"])
def generate():
    pass

if __name__ == "__main__":
    app.run(host="0.0.0.0",  port=int(os.environ.get("PORT", 5000)))