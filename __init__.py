from flask import Flask, render_template, request, send_from_directory, send_file, redirect, url_for, session
import os
import cv2
import easyocr
from openpyxl import Workbook
from openpyxl.styles import Alignment
import numpy as np
import sqlite3
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "instance", "database.db")
   
def create_app():
    app = Flask(__name__)
    app.secret_key = "secretkey"

 # ---------------- Admin Credentials ----------------
    ADMIN_USERNAME = "admin"
    ADMIN_PASSWORD = "password123"

    # ---------------- Upload Folder ----------------
    UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

    # ---------------- EasyOCR Reader ----------------
    reader = easyocr.Reader(['en'])

    # ---------------- Home ----------------
    @app.route("/")
    def home():
        return render_template("index.html")

    # ---------------- User Register ----------------
    @app.route("/register", methods=["GET", "POST"])
    def register():

        if request.method == "POST":

            username = request.form.get("username")
            email = request.form.get("email")
            password = request.form.get("password")

            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()

        # check if user already exists
            cur.execute("SELECT * FROM users WHERE username=?", (username,))
            existing_user = cur.fetchone()

            if existing_user:
                conn.close()
                return render_template("register.html", error="Username already exists")

        # insert new user
            cur.execute(
                "INSERT INTO users (username,email,password) VALUES (?,?,?)",
                (username, email, password)
            )

            conn.commit()
            conn.close()

            session['user_logged_in'] = True
            session['username'] = username

            return redirect(url_for("home"))

        return render_template("register.html")

    # ---------------- User Login ----------------
    @app.route("/login", methods=["GET", "POST"])
    def login():

        if request.method == "POST":

            username = request.form.get("username")
            password = request.form.get("password")

            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()

            cur.execute(
                "SELECT * FROM users WHERE username=? AND password=?",
                (username, password)
            )

            user = cur.fetchone()

            conn.close()

            if user:

                session['user_logged_in'] = True
                session['username'] = username
                session['user_id'] = user[0]

                return redirect(url_for("home"))

            else:
                return render_template("login.html", error="Invalid credentials")

        return render_template("login.html")

    # ---------------- User Logout ----------------
    @app.route("/logout")
    def logout():
        session.pop('user_logged_in', None)
        session.pop('username', None)
        return redirect(url_for("home"))

    # ---------------- About / Contact ----------------
    @app.route("/about")
    def about():
        return render_template("about.html")
    
    @app.route("/contact", methods=["GET","POST"])
    def contact():

        if request.method == "POST":

            name = request.form.get("name")
            email = request.form.get("email")
            subject = request.form.get("subject")
            message = request.form.get("message")

            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()

            cur.execute(
                "INSERT INTO contact (name,email,subject,message) VALUES (?,?,?,?)",
                (name,email,subject,message)
            )

            conn.commit()
            conn.close()

            return redirect(url_for("thank_you"))

        return render_template("contact.html")

    # ---------------- Upload ----------------
    @app.route("/upload", methods=["GET", "POST"])
    def upload():
        if not session.get('user_logged_in'):
            return redirect(url_for('login'))

        if request.method == "POST":
            file = request.files.get("file")
            if file and file.filename != "":
                filename = file.filename
                #image_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                #file.save(image_path)
                username = session.get("username")
                user_folder = os.path.join(app.config["UPLOAD_FOLDER"], username)
                os.makedirs(user_folder, exist_ok=True)

                image_path = os.path.join(user_folder, filename)
                file.save(image_path)
                # ---------------- Image Preprocessing ----------------
                img = cv2.imread(image_path)
                gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                gray = cv2.bitwise_not(gray)
                _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)

                # Save preprocessed image
                cv2.imwrite(image_path, thresh)

                # ---------------- OCR ----------------
                result = reader.readtext(image_path)

                # ---------------- Table Recognition ----------------
                # Collect all OCR words with positions
                words = []
                for bbox, text, conf in result:
                    x_min = min(pt[0] for pt in bbox)
                    x_max = max(pt[0] for pt in bbox)
                    y_min = min(pt[1] for pt in bbox)
                    y_max = max(pt[1] for pt in bbox)
                    words.append({
                        'text': text,
                        'x_min': x_min,
                        'x_max': x_max,
                        'y_min': y_min,
                        'y_max': y_max,
                        'y_center': (y_min + y_max)/2
                    })

                # ---------------- Cluster rows ----------------
                words.sort(key=lambda w: w['y_center'])
                rows = []
                current_row = []
                row_thresh = 15  # pixels
                for w in words:
                    if not current_row:
                        current_row.append(w)
                        continue
                    if abs(w['y_center'] - current_row[-1]['y_center']) <= row_thresh:
                        current_row.append(w)
                    else:
                        rows.append(current_row)
                        current_row = [w]
                if current_row:
                    rows.append(current_row)

                # ---------------- Sort columns within rows ----------------
                table_data = []
                col_gap_thresh = 20
                for r in rows:
                    r.sort(key=lambda w: w['x_min'])
                    row_cells = []
                    cell_text = ""
                    prev_x_max = None
                    for w in r:
                        if prev_x_max is None:
                            cell_text = w['text']
                        else:
                            if w['x_min'] - prev_x_max < col_gap_thresh:
                                cell_text += " " + w['text']
                            else:
                                row_cells.append(cell_text)
                                cell_text = w['text']
                        prev_x_max = w['x_max']
                    row_cells.append(cell_text)
                    table_data.append(row_cells)

                # ---------------- Create Excel ----------------
                excel_filename = f"{os.path.splitext(filename)[0]}.xlsx"
                excel_path = os.path.join(app.config["UPLOAD_FOLDER"], excel_filename)
                wb = Workbook()
                ws = wb.active
                ws.title = "Digitized Table"

                for row in table_data:
                    ws.append(row)

                for row_cells in ws.iter_rows():
                    for cell in row_cells:
                        cell.alignment = Alignment(wrap_text=True, vertical='top')

                # Auto adjust column width
                for col_cells in ws.columns:
                    max_len = max(len(str(c.value)) if c.value else 0 for c in col_cells)
                    ws.column_dimensions[col_cells[0].column_letter].width = max_len + 2

                wb.save(excel_path)

                return render_template(
                    "result.html",
                    image_filename=filename,
                    excel_filename=excel_filename,
                    table_data=table_data
                )

        return render_template("upload.html")

    # ---------------- Feedback ----------------
    @app.route("/feedback", methods=["GET", "POST"])
    def feedback():

        if request.method == "POST":

            name = request.form.get("name")
            email = request.form.get("email")
            message = request.form.get("feedback")

            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()

            cur.execute(
                "INSERT INTO feedback (name,email,message) VALUES (?,?,?)",
                (name,email,message)
            )

            conn.commit()
            conn.close()

            return redirect(url_for("thank_you"))

        return render_template("feedback.html")

    @app.route("/thank-you")
    def thank_you():
        return render_template("thank_you.html")

    # ---------------- Admin Login ----------------
    @app.route("/admin/login", methods=["GET", "POST"])
    def admin_login():
        if request.method == "POST":
            username = request.form.get("username")
            password = request.form.get("password")
            if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
                session["admin_logged_in"] = True
                return redirect(url_for("admin"))
            else:
                return render_template("admin_login.html", error="Invalid credentials")
        return render_template("admin_login.html")

    # ---------------- Admin Logout ----------------
    @app.route("/admin/logout")
    def admin_logout():
        session.pop("admin_logged_in", None)
        return redirect(url_for("admin_login"))

    # ---------------- Admin Panel ----------------
    @app.route("/admin")
    def admin():
        if not session.get("admin_logged_in"):
            return redirect(url_for("admin_login"))

        files = os.listdir(app.config["UPLOAD_FOLDER"])
        images = [f for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
        excels = [f for f in files if f.lower().endswith('.xlsx')]
        data = []
        for img in images:
            base = os.path.splitext(img)[0]
            matching_excel = next((e for e in excels if os.path.splitext(e)[0] == base), None)
            data.append({'image': img, 'excel': matching_excel})
        return render_template("admin.html", data=data)

    # ---------------- View / Download ----------------
    @app.route("/uploads/<filename>")
    def uploaded_file(filename):
        return send_from_directory(app.config["UPLOAD_FOLDER"], filename)


    @app.route("/delete_user/<int:user_id>")
    def delete_user(user_id):

        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()

        cur.execute("DELETE FROM users WHERE id=?", (user_id,))

        conn.commit()
        conn.close()

        return redirect(url_for("manage_users"))
    

    @app.route("/manage_users")
    def manage_users():
        if not session.get("admin_logged_in"):
            return redirect(url_for("admin_login"))


        #BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        #DB_PATH = os.path.join(BASE_DIR, "instance", "database.db")

        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()

        cur.execute("SELECT id, username, email, password FROM users")
        users = cur.fetchall()

        conn.close()

        return render_template("manage_users.html", users=users)
    @app.route("/manage_contact")
    def manage_contact():
        if not session.get("admin_logged_in"):
            return redirect(url_for("admin_login"))

        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()

        cur.execute("SELECT name, email, subject, message, created_at FROM contact")
        contacts = cur.fetchall()

        conn.close()

        return render_template("manage_contact.html", contacts=contacts)

    @app.route("/manage_feedback")
    def manage_feedback():
        if not session.get("admin_logged_in"):
            return redirect(url_for("admin_login"))

        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        
        cur.execute("SELECT name, email, message, created_at FROM feedback")
        feedbacks = cur.fetchall()

        conn.close()

        return render_template("manage_feedback.html", feedbacks=feedbacks)
    
    @app.route("/history")
    def history():
        if not session.get("user_logged_in"):
            return redirect(url_for("login"))

        username = session.get("username")
        user_folder = os.path.join(app.config["UPLOAD_FOLDER"], username)
        os.makedirs(user_folder, exist_ok=True)

        history_list = []

        files = os.listdir(user_folder)
        images = [f for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
        excels = [f for f in files if f.lower().endswith('.xlsx')]

        for img in images:
            base = os.path.splitext(img)[0]
            excel = next((e for e in excels if os.path.splitext(e)[0] == base), None)
            history_list.append({
                "date": "Uploaded File",
                "doc_name": img,
                "status": "Processed",
                "details": excel if excel else "No Excel Generated"
            })

        return render_template("history.html", user_history=history_list)
    @app.route("/download/<filename>")
    def download_file(filename):
        return send_file(os.path.join(app.config["UPLOAD_FOLDER"], filename), as_attachment=True)


    @app.route("/admin/history")
    def admin_history():
        if not session.get("admin_logged_in"):
            return redirect(url_for("admin_login"))

        history_list = []

    # Loop over all user folders
        for username in os.listdir(app.config["UPLOAD_FOLDER"]):
            user_folder = os.path.join(app.config["UPLOAD_FOLDER"], username)
            if not os.path.isdir(user_folder):
                continue

            files = os.listdir(user_folder)
            images = [f for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            excels = [f for f in files if f.lower().endswith('.xlsx')]

            for img in images:
                base = os.path.splitext(img)[0]
                excel = next((e for e in excels if os.path.splitext(e)[0] == base), None)
                history_list.append({
                    "user": username,
                    "date": "Uploaded File",
                    "doc_name": img,
                    "status": "Processed",
                    "details": excel if excel else "No Excel Generated"
                })

        return render_template("admin_history.html", all_history=history_list)
    return app

# ---------------- Run App ----------------
if __name__ == "__main__":
    app = create_app()
    app.run(debug=True)