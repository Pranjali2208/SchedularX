from flask import Flask, render_template, request, jsonify, send_file, session
from datetime import datetime, timedelta
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import os
import json
import secrets

from scheduler import create_schedule


# =========================================================
# FLASK APPLICATION / SECURITY
# =========================================================

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB upload limit
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(minutes=30)

instance_dir = os.path.join(app.root_path, "instance")
os.makedirs(instance_dir, exist_ok=True)
secret_file = os.path.join(instance_dir, "secret_key.txt")
if os.path.exists(secret_file):
    with open(secret_file, "r", encoding="utf-8") as f:
        app.secret_key = f.read().strip()
else:
    app.secret_key = secrets.token_hex(32)
    with open(secret_file, "w", encoding="utf-8") as f:
        f.write(app.secret_key)

AUTH_FILE = os.path.join(instance_dir, "admin_credentials.json")
DEFAULT_USERNAME = "admin"
DEFAULT_PASSWORD = "admin"
DEFAULT_RECOVERY_ANSWER = "scheduler"


def load_credentials():
    if not os.path.exists(AUTH_FILE):
        data = {
            "username": DEFAULT_USERNAME,
            "password_hash": generate_password_hash(DEFAULT_PASSWORD),
            "recovery_answer_hash": generate_password_hash(DEFAULT_RECOVERY_ANSWER),
        }
        save_credentials(data)
        return data
    try:
        with open(AUTH_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not all(k in data for k in ("username", "password_hash", "recovery_answer_hash")):
            raise ValueError("Invalid credentials file")
        return data
    except Exception:
        data = {
            "username": DEFAULT_USERNAME,
            "password_hash": generate_password_hash(DEFAULT_PASSWORD),
            "recovery_answer_hash": generate_password_hash(DEFAULT_RECOVERY_ANSWER),
        }
        save_credentials(data)
        return data


def save_credentials(data):
    temp = AUTH_FILE + ".tmp"
    with open(temp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    os.replace(temp, AUTH_FILE)


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("authenticated"):
            if request.path.startswith("/download/") or request.path.startswith("/generate"):
                return jsonify({"success": False, "error": "Your session has expired. Please log in again."}), 401
            return jsonify({"success": False, "error": "Authentication required."}), 401
        session.permanent = True
        return view(*args, **kwargs)
    return wrapped


# =========================================================
# AUTHENTICATION
# =========================================================

@app.route("/auth/login", methods=["POST"])
def auth_login():
    data = request.get_json(silent=True) or request.form
    username = str(data.get("username", "")).strip()
    password = str(data.get("password", ""))
    credentials = load_credentials()

    if username == credentials["username"] and check_password_hash(credentials["password_hash"], password):
        session.clear()
        session.permanent = True
        session["authenticated"] = True
        session["username"] = credentials["username"]
        return jsonify({"success": True, "username": credentials["username"]})

    return jsonify({"success": False, "error": "Invalid username or password."}), 401


@app.route("/auth/session", methods=["GET"])
def auth_session():
    return jsonify({
        "authenticated": bool(session.get("authenticated")),
        "username": session.get("username", "")
    })


@app.route("/auth/logout", methods=["POST"])
def auth_logout():
    session.clear()
    return jsonify({"success": True})


@app.route("/auth/reset-password", methods=["POST"])
def auth_reset_password():
    data = request.get_json(silent=True) or request.form
    username = str(data.get("username", "")).strip()
    recovery_answer = str(data.get("recovery_answer", "")).strip().lower()
    new_password = str(data.get("new_password", ""))
    confirm_password = str(data.get("confirm_password", ""))
    credentials = load_credentials()

    if username != credentials["username"]:
        return jsonify({"success": False, "error": "Username or recovery answer is incorrect."}), 400
    if not check_password_hash(credentials["recovery_answer_hash"], recovery_answer):
        return jsonify({"success": False, "error": "Username or recovery answer is incorrect."}), 400
    if len(new_password) < 6:
        return jsonify({"success": False, "error": "New password must contain at least 6 characters."}), 400
    if new_password != confirm_password:
        return jsonify({"success": False, "error": "New password and confirm password do not match."}), 400

    credentials["password_hash"] = generate_password_hash(new_password)
    save_credentials(credentials)
    session.clear()
    return jsonify({"success": True, "message": "Password reset successfully. Please log in with your new password."})


# =========================================================
# HOME PAGE
# =========================================================

@app.route("/")
def home():
    return render_template("index.html")


# =========================================================
# GENERATE RE-EXAM TIMETABLE
# =========================================================

@app.route("/generate", methods=["POST"])
@login_required
def generate():
    try:
        file = request.files.get("file")
        if not file:
            return jsonify({"success": False, "error": "Please upload an Excel file."}), 400
        if file.filename == "":
            return jsonify({"success": False, "error": "Please select a valid Excel file."}), 400

        original_name = secure_filename(file.filename)
        if not original_name.lower().endswith((".xlsx", ".xls")):
            return jsonify({"success": False, "error": "Only Excel files (.xlsx or .xls) are supported."}), 400

        upload_folder = os.path.join(app.root_path, "uploads")
        os.makedirs(upload_folder, exist_ok=True)
        file_path = os.path.join(upload_folder, original_name)
        file.save(file_path)

        start_date_text = request.form.get("start_date", "2026-10-26")
        if not start_date_text:
            return jsonify({"success": False, "error": "Please select an exam start date."}), 400
        try:
            start_date = datetime.strptime(start_date_text, "%Y-%m-%d").date()
        except ValueError:
            return jsonify({"success": False, "error": "Invalid date format. Please use YYYY-MM-DD."}), 400

        selected_course = request.form.get("course", "").strip()
        selected_year = request.form.get("year", "").strip()
        if not selected_course:
            return jsonify({"success": False, "error": "Please select a course."}), 400
        if not selected_year:
            return jsonify({"success": False, "error": "Please select an academic year."}), 400

        mapping_text = request.form.get("course_year_start_dates", "{}")
        try:
            course_year_start_dates = json.loads(mapping_text)
            if not isinstance(course_year_start_dates, dict):
                raise ValueError
        except (ValueError, TypeError, json.JSONDecodeError):
            return jsonify({"success": False, "error": "Invalid course/year start-date configuration."}), 400

        selected_key = f"{selected_course}|{selected_year}"
        mapped_start_date = course_year_start_dates.get(selected_key)
        if not mapped_start_date:
            return jsonify({"success": False, "error": f"Please add a starting date for {selected_course} - {selected_year} before generating."}), 400
        try:
            start_date = datetime.strptime(str(mapped_start_date).strip(), "%Y-%m-%d").date()
        except ValueError:
            return jsonify({"success": False, "error": f"Invalid start date for {selected_course} - {selected_year}."}), 400

        try:
            morning_papers = int(request.form.get("morning_papers", "5"))
            afternoon_papers = int(request.form.get("afternoon_papers", "4"))
        except (ValueError, TypeError):
            return jsonify({"success": False, "error": "Morning and Afternoon paper counts must be numbers."}), 400
        if morning_papers <= 0 or afternoon_papers <= 0:
            return jsonify({"success": False, "error": "Morning and Afternoon papers must be greater than 0."}), 400

        remove_labs = str(request.form.get("remove_labs", "true")).strip().lower() in ("true", "1", "yes", "on")

        output_folder = os.path.join(app.root_path, "generated_schedules")
        os.makedirs(output_folder, exist_ok=True)

        holiday_dates_text = request.form.get("holiday_dates", "[]")
        try:
            holiday_dates = json.loads(holiday_dates_text)
            if not isinstance(holiday_dates, list):
                raise ValueError
        except (ValueError, TypeError, json.JSONDecodeError):
            return jsonify({"success": False, "error": "Invalid holiday dates."}), 400

        validated_holidays = []
        for holiday_text in holiday_dates:
            try:
                holiday_date = datetime.strptime(str(holiday_text), "%Y-%m-%d").date()
            except ValueError:
                return jsonify({"success": False, "error": f"Invalid holiday date: {holiday_text}"}), 400
            if holiday_date not in validated_holidays:
                validated_holidays.append(holiday_date)

        result = create_schedule(
            excel_path=file_path,
            start_date=start_date,
            morning_capacity=morning_papers,
            afternoon_capacity=afternoon_papers,
            output_folder=output_folder,
            holidays=validated_holidays,
            remove_labs=remove_labs
        )

        return jsonify({
            "success": True,
            "students": result["students"],
            "subjects": result["subjects"],
            "exam_days": result["exam_days"],
            "schedule": result["schedule"],
            "cleaned_file": result["cleaned_file"],
            "timetable_file": result["timetable_file"],
            "morning_capacity": morning_papers,
            "afternoon_capacity": afternoon_papers,
            "remove_labs": remove_labs,
            "course": selected_course,
            "year": selected_year,
            "start_date": start_date.strftime("%Y-%m-%d"),
            "course_year_start_dates": course_year_start_dates,
            "holidays": [d.strftime("%Y-%m-%d") for d in validated_holidays]
        })
    except Exception as e:
        print("\nERROR:", str(e))
        return jsonify({"success": False, "error": str(e)}), 500


# =========================================================
# DOWNLOADS
# =========================================================

@app.route("/download/timetable")
@login_required
def download_timetable():
    try:
        file_path = os.path.join(app.root_path, "generated_schedules", "generated_reexam_timetable.xlsx")
        if not os.path.exists(file_path):
            return "Timetable file not found. Please generate the timetable first.", 404
        return send_file(file_path, as_attachment=True, download_name="generated_reexam_timetable.xlsx")
    except Exception as e:
        return str(e), 500


@app.route("/download/cleaned")
@login_required
def download_cleaned():
    try:
        file_path = os.path.join(app.root_path, "generated_schedules", "cleaned_reexam_database.xlsx")
        if not os.path.exists(file_path):
            return "Cleaned database file not found. Please generate the timetable first.", 404
        return send_file(file_path, as_attachment=True, download_name="cleaned_reexam_database.xlsx")
    except Exception as e:
        return str(e), 500


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
