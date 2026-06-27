from flask import Flask, render_template, request, redirect, session, send_file, after_this_request
from docxtpl import DocxTemplate, InlineImage
from werkzeug.utils import secure_filename #for making the file name secure like my file.csv -> my_file.csv automatically
from PIL import Image
from docx.shared import Mm
import tempfile
import pandas as pd
import os
import json
import csv


import shutil
# import dotenv

app = Flask(__name__)
# dotenv.load_dotenv()
app.secret_key = os.getenv("SECRETKEY")


#resizeing the image
def resize_keep_ratio(img, target_width):
    w, h = img.size

    if w <= target_width:
        return img

    new_height = int(h * target_width / w)

    return img.resize(
        (target_width, new_height),
        Image.LANCZOS
    )



@app.route("/")
def index():
    return render_template("index.html")

@app.before_request
def require_login():

    allowed_routes = ["login"]

    if request.endpoint in allowed_routes:
        return

    if not session.get("logged_in"):
        return redirect("/login")


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
    csv_path = os.path.join(UPLOAD_FOLDER, filename)          # it join the temp folder path with the secured file name
    
    
    try:
        df = pd.read_csv(csv_file)
    except Exception as e:
        return f"❌ Error reading CSV: {e}"

    required_columns = ["Name", "RollNo"]

    for col in required_columns:
        if col not in df.columns:
            return f"❌ Missing required column: {col}"

    df = df[required_columns]
    df.to_csv(csv_path)
    del df
    # reader = csv.reader(csv_file.stream.read().decode("utf-8").splitlines()) #initialize the csv reader

    # header = next(reader)

    # required = {"Name", "RollNo"}
    # missing = required - set(header)
    
    
    # csv_file.save(csv_path)
    # if missing:
    #     return f"❌ Missing columns: {', '.join(missing)}"  
          
    
    #save the file into the temp folder path
    session["csv_path"] = csv_path



    # =========================
    # Store in Session
    # =========================f
    session["event_details"] = {
        "code": code,
        "date": date,
        "duration": duration,
        "venue": venue,
        "register1": register1,
        "register2": register2,
        "register3": register3,

    }

    # =========================
    # Debug Print
    # =========================
    # print("\n=== EVENT DETAILS ===")
    # print(session["event_details"])

    # print("\n=== PARTICIPANTS ===")
    # for p in participants:
    #     print(p)

    return redirect("/event")
@app.route("/event", methods= ["GET","POST"])
def image():
    if request.method == "POST":
        UPLOAD_FOLDER = session.get("UPLOAD_FOLDER")
        imagelist = []
        for a in range(1,5):
            file = request.files.get(f"image{a}")

            if file and file.filename != "":

                # Get original extension
                ext = os.path.splitext(file.filename)[1].lower()

                filename = f"image{a}{ext}"
                path = os.path.join(UPLOAD_FOLDER, filename)    #to merge the path of the temp folder with the image name

                # modify and correcting the format of the inserted image 
#                 img = Image.open(file.stream)
#                 if img.mode != "RGB":
#                     img = img.convert("RGB")
                
#                 img.save(
#     path,
#     format="JPEG",
#     quality=75,
#     optimize=False
# )
                file.save(path)
                # img = resize_keep_ratio(img, 1200)
                # img.save(path, quality = 80)
                # print(f"image done")


                imagelist.append(path)
        session["imagespath"] = imagelist
        return redirect("/teams")


    return render_template("image.html")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":

        password = request.form.get("password")

        if password == os.getenv("password"):
            session["logged_in"] = True
            return redirect("/")

        return "Wrong password"

    return """
    <form method="POST">
        <h2>Welcome To Report Maker <h2>
        <h2>Enter Password</h2>
        <input type="password" name="password" required>
        <button type="submit">Login</button>
    </form>
    """

@app.route("/teams")
def teams():
        # =========================
    # Read CSV
    # =========================
    csv_path = session.get("csv_path")
    try:
        df = pd.read_csv(csv_path)
    except Exception as e:
        return f"❌ Error reading CSV: {e}"

    # =========================
    # Convert to List
    # =========================
    participants = df.to_dict(orient="records")

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
            "conclusion": team["conclusion"],
            "mvp": image_obj
        })

    imagelist = session.get("imagespath")
    print(imagelist)

    for a, image  in enumerate(imagelist):
            context[f"image{a+1}"] = InlineImage(doc, image, width=Mm(82))
  



    # for debbuging
    # print("\n=== FINAL CONTEXT ===")
    # print(json.dumps(context, indent=4))


    # 1. Get the safe data from the user
    # print(context)

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