import os
import sys
import sqlite3
import smtplib
import urllib.request  # <--- add this
from email.message import EmailMessage
from datetime import datetime


from flask import (
    Flask,
    request,
    jsonify,
    send_file,
    abort,
    render_template_string,
    url_for,
)

# PDF support
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import (
        SimpleDocTemplate,
        Table,
        TableStyle,
        Paragraph,
        Spacer,
        Image as RLImage,
        LongTable,
    )
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch

    REPORTLAB_INSTALLED = True
except ImportError:
    REPORTLAB_INSTALLED = False

PRIMARY_COLOR = "#1f618d"
SECONDARY_COLOR = "#2e86c1"
ACCENT_COLOR = "#27ae60"
DANGER_COLOR = "#c0392b"

COLLEGE_LOGO_PATH = "assets/images/logoclg.png"
TN_LOGO_PATH = "assets/images/tn_logo.png"

DEPARTMENTS = [
    "B.A. Economics",
    "B.A. English",
    "B.A. Tamil",
    "B.Sc. Botany",
    "B.Sc. Chemistry",
    "B.Sc. Mathematics",
    "B.Sc. Physics",
    "B.Sc. Computer Science",
    "B.Sc. Marine Biology",
    "B.Com.",
    "B.Com. (CA)",
    "Other",
]

YEARS = [1, 2, 3]
SEMESTERS = [1, 2, 3, 4, 5, 6]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


class ReportCardSystem:
    def __init__(self):
        self.data_dir = "student_records"
        self.reports_dir = os.path.join(self.data_dir, "report_cards_pdf")
        self.backup_dir = os.path.join(self.data_dir, "backups")
        self.db_file = os.path.join(self.data_dir, "students.db")

        os.makedirs(self.data_dir, exist_ok=True)
        os.makedirs(self.reports_dir, exist_ok=True)
        os.makedirs(self.backup_dir, exist_ok=True)

        self.init_database()

    def init_database(self):
        conn = None
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()

            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS departments (
                    dept_code TEXT PRIMARY KEY,
                    dept_name TEXT NOT NULL UNIQUE
                )
                """
            )

            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS students (
                    student_id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    dept_code TEXT NOT NULL,
                    year INTEGER NOT NULL,
                    semester INTEGER NOT NULL,
                    email TEXT,
                    program_type TEXT,
                    duration_years INTEGER,
                    course_name TEXT,
                    created_date TEXT,
                    FOREIGN KEY (dept_code) REFERENCES departments(dept_code)
                )
                """
            )

            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS subjects (
                    subject_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    subject_name TEXT NOT NULL,
                    subject_code TEXT NOT NULL UNIQUE,
                    year INTEGER NOT NULL,
                    semester INTEGER NOT NULL,
                    max_marks INTEGER DEFAULT 100
                )
                """
            )

            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS marks (
                    mark_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student_id TEXT NOT NULL,
                    subject_code TEXT NOT NULL,
                    marks REAL NOT NULL,
                    FOREIGN KEY (student_id) REFERENCES students(student_id)
                        ON DELETE CASCADE,
                    FOREIGN KEY (subject_code) REFERENCES subjects(subject_code)
                        ON DELETE CASCADE,
                    UNIQUE(student_id, subject_code)
                )
                """
            )

            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS attendance (
                    att_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student_id TEXT NOT NULL,
                    month TEXT NOT NULL,
                    classes_held INTEGER NOT NULL,
                    classes_attended INTEGER NOT NULL,
                    FOREIGN KEY (student_id) REFERENCES students(student_id)
                        ON DELETE CASCADE,
                    UNIQUE(student_id, month)
                )
                """
            )

            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS settings (
                    id INTEGER PRIMARY KEY CHECK (id = 1),
                    smtp_email TEXT,
                    smtp_app_password TEXT
                )
                """
            )
            cursor.execute("INSERT OR IGNORE INTO settings (id) VALUES (1)")

            conn.commit()

            if cursor.execute("SELECT COUNT(*) FROM departments").fetchone()[0] == 0:
                self._populate_default_departments(conn)

        except sqlite3.Error as e:
            print(f"Database initialization error: {e}")
            sys.exit(1)
        finally:
            if conn:
                conn.close()

    def _populate_default_departments(self, conn):
        cursor = conn.cursor()
        for i, dept in enumerate(DEPARTMENTS):
            code = dept[:3].upper() + str(i + 1).zfill(2)
            cursor.execute(
                "INSERT INTO departments (dept_code, dept_name) VALUES (?, ?)",
                (code, dept),
            )
        conn.commit()

    def get_all_departments(self):
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            cursor.execute(
                "SELECT dept_code, dept_name FROM departments ORDER BY dept_name"
            )
            rows = cursor.fetchall()
            conn.close()
            return rows
        except sqlite3.Error:
            return []

    def add_or_update_student(
        self,
        student_id,
        name,
        dept_name,
        year,
        semester,
        email,
        program_type,
        duration_years,
        course_name,
        marks_data,
        attendance_data,
    ):
        conn = None
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()

            dept_code = next(
                (d[0] for d in self.get_all_departments() if d[1] == dept_name),
                "OTH99",
            )

            cursor.execute(
                """
                INSERT OR REPLACE INTO students
                (student_id, name, dept_code, year, semester, email,
                 program_type, duration_years, course_name, created_date)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    student_id,
                    name,
                    dept_code,
                    year,
                    semester,
                    email,
                    program_type,
                    duration_years,
                    course_name,
                    datetime.now().strftime("%Y-%m-%d"),
                ),
            )

            cursor.execute("DELETE FROM marks WHERE student_id = ?", (student_id,))
            cursor.execute(
                "DELETE FROM attendance WHERE student_id = ?", (student_id,)
            )

            for code, marks in marks_data.items():
                if marks is not None and str(marks).strip() != "":
                    cursor.execute(
                        """
                        INSERT INTO marks (student_id, subject_code, marks)
                        VALUES (?, ?, ?)
                        """,
                        (student_id, code, float(marks)),
                    )

            for month, data in attendance_data.items():
                held = int(data.get("held", 0))
                attended = int(data.get("attended", 0))
                if held > 0:
                    cursor.execute(
                        """
                        INSERT INTO attendance (student_id, month,
                                                classes_held, classes_attended)
                        VALUES (?, ?, ?, ?)
                        """,
                        (student_id, month, held, attended),
                    )

            conn.commit()
            return True, f"Student {student_id} data saved/updated successfully."
        except sqlite3.IntegrityError:
            return False, f"Error: Student ID {student_id} already exists or data is corrupt."
        except Exception as e:
            return False, f"Database error during save: {e}"
        finally:
            if conn:
                conn.close()

    def add_subject(self, name, code, year, semester, max_marks):
        conn = None
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            cursor.execute(
                """
                INSERT INTO subjects
                (subject_name, subject_code, year, semester, max_marks)
                VALUES (?, ?, ?, ?, ?)
                """,
                (name, code, year, semester, max_marks),
            )
            conn.commit()
            return True, f"Subject '{name}' added successfully."
        except sqlite3.IntegrityError:
            return False, f"Subject code '{code}' already exists."
        except sqlite3.Error as e:
            return False, f"Database error: {e}"
        finally:
            if conn:
                conn.close()

    def get_all_subjects(self):
        conn = None
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT subject_code, subject_name, year, semester, max_marks
                FROM subjects
                ORDER BY year, semester, subject_name
                """
            )
            return cursor.fetchall()
        except sqlite3.Error:
            return []
        finally:
            if conn:
                conn.close()

    def get_student_details_and_data(self, student_id):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        student_row = cursor.execute(
            """
            SELECT s.student_id, s.name, s.dept_code, s.year, s.semester,
                   s.email, s.program_type, s.duration_years, s.course_name,
                   d.dept_name
            FROM students s
            JOIN departments d ON s.dept_code = d.dept_code
            WHERE s.student_id = ?
            """,
            (student_id,),
        ).fetchone()
        if not student_row:
            conn.close()
            return None, None, None

        student = dict(
            zip(
                [
                    "student_id",
                    "name",
                    "dept_code",
                    "year",
                    "semester",
                    "email",
                    "program_type",
                    "duration_years",
                    "course_name",
                    "dept_name",
                ],
                student_row,
            )
        )

        marks = cursor.execute(
            """
            SELECT s.subject_code, s.subject_name, s.max_marks, m.marks
            FROM marks m
            JOIN subjects s ON m.subject_code = s.subject_code
            WHERE m.student_id = ?
            """,
            (student_id,),
        ).fetchall()
        marks_data = [
            dict(zip(["subject_code", "subject_name", "max_marks", "marks"], row))
            for row in marks
        ]

        attendance = cursor.execute(
            """
            SELECT month, classes_held, classes_attended
            FROM attendance
            WHERE student_id = ?
            ORDER BY CASE month
                WHEN 'Jan' THEN 1 WHEN 'Feb' THEN 2 WHEN 'Mar' THEN 3
                WHEN 'Apr' THEN 4 WHEN 'May' THEN 5 WHEN 'Jun' THEN 6
                WHEN 'Jul' THEN 7 WHEN 'Aug' THEN 8 WHEN 'Sep' THEN 9
                WHEN 'Oct' THEN 10 WHEN 'Nov' THEN 11 WHEN 'Dec' THEN 12
            END
            """,
            (student_id,),
        ).fetchall()
        attendance_data = [
            dict(zip(["month", "classes_held", "classes_attended"], row))
            for row in attendance
        ]

        conn.close()
        return student, marks_data, attendance_data

    def get_all_students_summary(self):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        students = cursor.execute(
            """
            SELECT s.student_id, s.name, d.dept_name, s.year, s.semester,
                   s.email, s.program_type, s.duration_years, s.course_name
            FROM students s
            JOIN departments d ON s.dept_code = d.dept_code
            ORDER BY s.dept_code, s.year, s.student_id
            """
        ).fetchall()

        all_summaries = []
        for student_row in students:
            student_id = student_row[0]
            summary = list(student_row)

            att_data = cursor.execute(
                """
                SELECT SUM(classes_held), SUM(classes_attended)
                FROM attendance WHERE student_id = ?
                """,
                (student_id,),
            ).fetchone()
            held, attended = att_data if att_data and att_data[0] is not None else (0, 0)
            att_percent = (attended / held * 100) if held > 0 else 0.0
            summary.append(f"{att_percent:.2f}%")

            mark_data = cursor.execute(
                """
                SELECT SUM(m.marks), SUM(s.max_marks)
                FROM marks m
                JOIN subjects s ON m.subject_code = s.subject_code
                WHERE m.student_id = ?
                """,
                (student_id,),
            ).fetchone()
            total_marks, total_max = (
                mark_data if mark_data and mark_data[0] is not None else (0, 0)
            )
            mark_percent = (total_marks / total_max * 100) if total_max > 0 else 0.0
            summary.append(f"{mark_percent:.2f}%")

            all_summaries.append(summary)

        conn.close()
        return all_summaries

    def calculate_grade(self, percentage):
        if percentage >= 90:
            return "O"
        elif percentage >= 80:
            return "A+"
        elif percentage >= 70:
            return "A"
        elif percentage >= 60:
            return "B+"
        elif percentage >= 50:
            return "B"
        elif percentage >= 40:
            return "C"
        else:
            return "F"

    def get_email_settings(self):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        row = cursor.execute(
            "SELECT smtp_email, smtp_app_password FROM settings WHERE id = 1"
        ).fetchone()
        conn.close()
        if not row:
            return None, None
        return row[0], row[1]

    def save_email_settings(self, smtp_email, smtp_app_password):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE settings SET smtp_email = ?, smtp_app_password = ? WHERE id = 1",
            (smtp_email, smtp_app_password),
        )
        conn.commit()
        conn.close()

    def email_report_to_student(self, student_id):
        smtp_email, smtp_app_password = self.get_email_settings()
        if not smtp_email or not smtp_app_password:
            return False, "Email settings not configured."

        student, _, _ = self.get_student_details_and_data(student_id)
        if not student:
            return False, f"Student {student_id} not found."

        if not student.get("email"):
            return False, "Student email is empty."

        if not REPORTLAB_INSTALLED:
            return False, "reportlab not installed; cannot generate PDF."

        pdf_path = self.generate_student_report_pdf(student_id)

        msg = EmailMessage()
        msg["Subject"] = (
            f"SGAC Report Card - {student['name']} ({student['student_id']})"
        )
        msg["From"] = smtp_email
        msg["To"] = student["email"]
        body = (
            f"Dear {student['name']},\n\n"
            "Please find attached your report card generated by the "
            "Sethupathy Government Arts College portal.\n\n"
            "This is a system‑generated email.\n\n"
            "Regards,\n"
            "Sethupathy Government Arts College"
        )
        msg.set_content(body)

        with open(pdf_path, "rb") as f:
            data = f.read()
        msg.add_attachment(
            data,
            maintype="application",
            subtype="pdf",
            filename=os.path.basename(pdf_path),
        )

        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login(smtp_email, smtp_app_password)
                server.send_message(msg)
        except Exception as e:
            return False, f"Error sending email: {e}"

        return True, "Report emailed successfully."

    def generate_student_report_pdf(self, student_id):
        if not REPORTLAB_INSTALLED:
            raise ImportError("reportlab not installed. Cannot generate PDF.")

        student, marks_data, attendance_data = self.get_student_details_and_data(
            student_id
        )
        if not student:
            raise ValueError(f"Student ID {student_id} not found.")

        file_path = os.path.join(
            self.reports_dir,
            f"{student_id}_{student['name'].replace(' ', '_')}_Report.pdf",
        )

        doc = SimpleDocTemplate(
            file_path,
            pagesize=A4,
            rightMargin=40,
            leftMargin=40,
            topMargin=60,
            bottomMargin=40,
        )
        story = []
        styles = getSampleStyleSheet()

        # --- fetch logos from official URLs into temp folder ---
                # --- fetch logos from official URLs into temp folder ---
        tn_logo_url = "https://sgacrmd.edu.in/assets/tn_logo.png"
        college_logo_url = "https://sgacrmd.edu.in/assets/logoclg.png"
        tmp_dir = os.path.join(self.data_dir, "tmp_logos")
        os.makedirs(tmp_dir, exist_ok=True)
        tn_logo_local = os.path.join(tmp_dir, "tn_logo.png")
        clg_logo_local = os.path.join(tmp_dir, "logoclg.png")

        # Try to download; if it fails, just set to None (no fallback files)
        try:
            if not os.path.exists(tn_logo_local):
                urllib.request.urlretrieve(tn_logo_url, tn_logo_local)
        except Exception:
            tn_logo_local = None

        try:
            if not os.path.exists(clg_logo_local):
                urllib.request.urlretrieve(college_logo_url, clg_logo_local)
        except Exception:
            clg_logo_local = None

        logo_width = 0.8 * inch
        logo_height = 0.8 * inch

        logo1 = (
            RLImage(tn_logo_local, width=logo_width, height=logo_height)
            if tn_logo_local and os.path.exists(tn_logo_local)
            else Paragraph("", styles["Normal"])
        )
        logo2 = (
            RLImage(clg_logo_local, width=logo_width, height=logo_height)
            if clg_logo_local and os.path.exists(clg_logo_local)
            else Paragraph("", styles["Normal"])
        )


        # --- colored college header band (matches website feel) ---
        header_table = Table(
            [
                [
                    logo1,
                    Paragraph(
                        "<b><font size=14 color='white'>SETHUPATHY GOVERNMENT ARTS COLLEGE</font></b><br/>"
                        "<font size=9 color='white'>Affiliated to Alagappa University | Accredited with 'A' Grade by NAAC</font><br/>"
                        "<font size=8 color='white'>Ramanathapuram - 623 502, Tamil Nadu, India</font>",
                        styles["Normal"],
                    ),
                    logo2,
                ]
            ],
            colWidths=[1.0 * inch, 4.7 * inch, 1.0 * inch],
        )
        header_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(PRIMARY_COLOR)),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (0, 0), (0, 0), "LEFT"),
                    ("ALIGN", (1, 0), (1, 0), "CENTER"),
                    ("ALIGN", (2, 0), (2, 0), "RIGHT"),
                    ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        story.append(header_table)
        story.append(Spacer(1, 0.2 * inch))

        title_para = Paragraph(
            "<b><font size=13>Comprehensive Student Performance Report</font></b>",
            styles["Title"],
        )
        story.append(title_para)
        story.append(Spacer(1, 0.15 * inch))

        # --- 1. Student profile block ---
        story.append(Paragraph("<b>1. Student Profile</b>", styles["h3"]))
        story.append(Spacer(1, 0.08 * inch))

        profile_data = [
            [
                Paragraph("<b>Student ID:</b>", styles["Normal"]),
                student["student_id"],
                Paragraph("<b>Name:</b>", styles["Normal"]),
                student["name"],
            ],
            [
                Paragraph("<b>Department:</b>", styles["Normal"]),
                student["dept_name"],
                Paragraph("<b>Year / Semester:</b>", styles["Normal"]),
                f"{student['year']} / {student['semester']}",
            ],
            [
                Paragraph("<b>Program / Duration:</b>", styles["Normal"]),
                f"{student.get('program_type', '')} / {student.get('duration_years', '')} years",
                Paragraph("<b>Course:</b>", styles["Normal"]),
                student.get("course_name", ""),
            ],
            [
                Paragraph("<b>Email:</b>", styles["Normal"]),
                student.get("email", ""),
                Paragraph("<b>Date Generated:</b>", styles["Normal"]),
                datetime.now().strftime("%Y-%m-%d"),
            ],
        ]
        prof_table = Table(
            profile_data,
            colWidths=[1.7 * inch, 2.3 * inch, 1.7 * inch, 2.3 * inch],
        )
        prof_table.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.6, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ]
            )
        )
        story.append(prof_table)
        story.append(Spacer(1, 0.3 * inch))

        # --- 2. Subject marks ---
        story.append(Paragraph("<b>2. Subject Marks Summary</b>", styles["h3"]))
        story.append(Spacer(1, 0.05 * inch))

        marks_table_data = [
            [
                Paragraph("<b>Subject Code</b>", styles["Normal"]),
                Paragraph("<b>Subject Name</b>", styles["Normal"]),
                Paragraph("<b>Max Marks</b>", styles["Normal"]),
                Paragraph("<b>Marks Obtained</b>", styles["Normal"]),
            ]
        ]
        total_marks = 0
        total_max_marks = 0
        for mark in marks_data:
            marks_table_data.append(
                [
                    mark["subject_code"],
                    Paragraph(mark["subject_name"], styles["Normal"]),
                    str(mark["max_marks"]),
                    str(mark["marks"]),
                ]
            )
            total_marks += mark["marks"]
            total_max_marks += mark["max_marks"]

        marks_table = LongTable(
            marks_table_data,
            colWidths=[1.1 * inch, 3.0 * inch, 1.3 * inch, 1.3 * inch],
            repeatRows=1,
        )
        marks_table.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.6, colors.grey),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(SECONDARY_COLOR)),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("ALIGN", (1, 1), (1, -1), "LEFT"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ]
            )
        )
        story.append(marks_table)

        percentage = (total_marks / total_max_marks * 100) if total_max_marks > 0 else 0
        grade = self.calculate_grade(percentage)

        story.append(Spacer(1, 0.2 * inch))
        story.append(
            Paragraph(
                f"<b>Overall Total:</b> {total_marks} / {total_max_marks}",
                styles["Normal"],
            )
        )
        story.append(
            Paragraph(
                f"<b>Percentage:</b> <font color='blue'>{percentage:.2f}%</font>",
                styles["Normal"],
            )
        )
        story.append(
            Paragraph(
                f"<b>Grade:</b> <font color='red'>{grade}</font>",
                styles["Normal"],
            )
        )
        story.append(Spacer(1, 0.3 * inch))

        # --- 3. Attendance ---
        story.append(Paragraph("<b>3. Monthly Attendance Record</b>", styles["h3"]))
        story.append(Spacer(1, 0.05 * inch))

        att_table_data = [
            [
                Paragraph("<b>Month</b>", styles["Normal"]),
                Paragraph("<b>Classes Held</b>", styles["Normal"]),
                Paragraph("<b>Classes Attended</b>", styles["Normal"]),
                Paragraph("<b>Percentage</b>", styles["Normal"]),
            ]
        ]
        total_att_held = 0
        total_att_attended = 0
        for att in attendance_data:
            held = att["classes_held"]
            attended = att["classes_attended"]
            att_pct = (attended / held * 100) if held > 0 else 0
            att_table_data.append(
                [att["month"], str(held), str(attended), f"{att_pct:.2f}%"]
            )
            total_att_held += held
            total_att_attended += attended

        att_table = LongTable(
            att_table_data,
            colWidths=[1.2 * inch, 1.4 * inch, 1.8 * inch, 1.6 * inch],
            repeatRows=1,
        )
        att_table.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.6, colors.grey),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(SECONDARY_COLOR)),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ]
            )
        )
        story.append(att_table)

        final_att_pct = (
            (total_att_attended / total_att_held * 100) if total_att_held > 0 else 0
        )
        att_color = DANGER_COLOR if final_att_pct < 75 else ACCENT_COLOR

        story.append(Spacer(1, 0.2 * inch))
        story.append(
            Paragraph(
                f"<b>Overall Attendance:</b> {total_att_attended} / {total_att_held} classes",
                styles["Normal"],
            )
        )
        story.append(
            Paragraph(
                f"<b>Final Percentage:</b> "
                f"<font color='{att_color}'>{final_att_pct:.2f}%</font> "
                "(75% is required)",
                styles["Normal"],
            )
        )
        story.append(Spacer(1, 0.3 * inch))

        # --- footer note ---
        story.append(Paragraph("<hr/>", styles["Normal"]))
        story.append(
            Paragraph(
                "<i>This is a system-generated report card. "
                "Head of Department signature is required for official use.</i>",
                styles["Normal"],
            )
        )

        doc.build(story)
        return file_path

app = Flask(__name__)
system = ReportCardSystem()

INDEX_HTML = """
<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>SGAC | Student Performance & Attendance Portal</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description"
          content="Internal Student Performance & Attendance Portal for Sethupathy Government Arts College, Ramanathapuram.">
    <link rel="stylesheet"
          href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap"
          rel="stylesheet">
    <style>
        :root {
            --primary-blue: #1f618d;
            --secondary-blue: #2e86c1;
            --accent-gold: #f1c40f;
            --bg-light: #f5f7fa;
            --text-dark: #333;
            --shadow-sm: 0 2px 4px rgba(0,0,0,0.05);
        }
        * { box-sizing:border-box; }
        body {
            margin:0;
            font-family:'Roboto',sans-serif;
            background:var(--bg-light);
            color:var(--text-dark);
        }
        a { text-decoration:none; color:inherit; }

        .top-bar {
            background:var(--primary-blue);
            color:#fff;
            font-size:0.9rem;
        }
        .top-bar-content {
            width:95%;
            max-width:1200px;
            margin:0 auto;
            display:flex;
            justify-content:space-between;
            align-items:center;
            padding:5px 0;
        }
        .top-bar a { color:#fff; margin-right:15px; }
        .btn-sm-top {
            padding:4px 10px;
            border-radius:20px;
            background:#fff;
            color:var(--primary-blue) !important;
            font-size:0.8rem;
            font-weight:500;
        }

        header {
            background:#fff;
            box-shadow:0 2px 4px rgba(0,0,0,0.05);
        }
        .header-container {
            width:95%;
            max-width:1200px;
            margin:0 26%;
            display:flex;
            justify-content:space-between;
            align-items:center;
            padding:10px 0;
        }
        .logo-section {
            display:flex;
            align-items:center;
            gap:10px;
        }
        .govt-logo, .college-logo {
            width:60px;
            height:60px;
            object-fit:contain;
        }
        .college-title h1 {
            margin:0;
            font-size:2.4rem;
            color:var(--primary-blue);
        }
        .college-title p {
            margin:2px 0;
            font-size:0.9rem;
            color:var(--secondary-blue);
        }
        .college-title span {
            font-size:0.8rem;
            color:#555;
        }

        .page-wrapper {
            width:95%;
            max-width:1200px;
            margin:25px auto 40px;
        }

        .hero-card {
            background:linear-gradient(135deg,#1f618d,#145374);
            color:#fff;
            border-radius:10px;
            padding:20px 24px 18px;
            box-shadow:var(--shadow-sm);
        }
        .hero-header {
            display:flex;
            flex-direction:column;
            align-items:center;
            text-align:center;
            gap:4px;
        }
        .hero-header h2 {
            margin:0;
            font-size:1.7rem;
        }
        .hero-header p {
            margin:0;
            font-size:0.95rem;
        }
        .hero-actions {
            margin-top:14px;
            display:flex;
            justify-content:center;
            gap:10px;
            flex-wrap:wrap;
            font-size:0.85rem;
        }
        .hero-pill {
            background:rgba(255,255,255,0.12);
            border-radius:20px;
            padding:6px 12px;
            display:flex;
            align-items:center;
            gap:6px;
        }

        .cards-grid {
            margin-top:22px;
            display:grid;
            grid-template-columns:repeat(auto-fit,minmax(260px,1fr));
            gap:16px;
        }
        .action-card {
            background:#fff;
            border-radius:8px;
            box-shadow:var(--shadow-sm);
            padding:16px 18px;
            border-top:4px solid var(--secondary-blue);
            display:flex;
            flex-direction:column;
            gap:6px;
        }
        .action-card h3 {
            margin:0;
            font-size:1.05rem;
            color:var(--primary-blue);
            display:flex;
            align-items:center;
            gap:8px;
        }
        .action-card p {
            margin:0;
            font-size:0.9rem;
            color:#555;
        }
        .action-card .btn {
            margin-top:8px;
            align-self:flex-start;
            padding:6px 13px;
            border-radius:20px;
            background:var(--primary-blue);
            color:#fff;
            font-size:0.85rem;
            font-weight:500;
        }

        footer {
            background:#111;
            color:#aaa;
            font-size:0.8rem;
            text-align:center;
            padding:8px 0;
        }
        footer a { color:var(--accent-gold); }

        @media (max-width:768px) {
            .header-container { flex-direction:column; align-items:flex-start; gap:8px; }
        }
    </style>
</head>
<body>
<div class="top-bar">
    <div class="top-bar-content">
        <div>
            <a href="tel:+914567220268"><i class="fas fa-phone-alt"></i> +91-4567-220268</a>
            <a href="mailto:principal@sgacrmd.edu.in"><i class="fas fa-envelope"></i> principal@sgacrmd.edu.in</a>
        </div>
        <div>
            <a href="https://sgacrmd.edu.in" target="_blank" class="btn-sm-top">College Website</a>
        </div>
    </div>
</div>

<header>
    <div class="header-container">
        <div class="logo-section">
            <img src="https://sgacrmd.edu.in/assets/tn_logo.png" alt="TN Govt Logo" class="govt-logo">
            <div class="college-title">
                <h1>Sethupathy Government Arts College</h1>
                <p>Internal Student Performance & Attendance Portal</p>
                <span>Ramanathapuram - 623 502, Tamil Nadu, India</span>
            </div>
            <img src="https://sgacrmd.edu.in/assets/logoclg.png" alt="College Logo" class="college-logo">
        </div>
    </div>
</header>

<div class="page-wrapper">
    <div class="hero-card">
        <div class="hero-header">
            <h2><i class="fas fa-chart-line"></i> Student Performance Dashboard</h2>
            <p>Manage students, subjects, marks, attendance and generate official PDF report cards with email delivery.</p>
        </div>
        <div class="hero-actions">
            <div class="hero-pill">
                <i class="fas fa-file-pdf"></i>
                <span>PDF Report</span>
            </div>
            <div class="hero-pill">
                <i class="fas fa-paper-plane"></i>
                <span>Email via Gmail App Password</span>
            </div>
            <div class="hero-pill">
                <i class="fas fa-database"></i>
                <span>SQLite Storage</span>
            </div>
        </div>
    </div>

    <div class="cards-grid">
        <div class="action-card">
            <h3><i class="fas fa-user-plus"></i> Add / Update Student</h3>
            <p>Create or update a student record with marks and monthly attendance.</p>
            <a href="{{ url_for('ui_new_student') }}" class="btn">Open Form</a>
        </div>

        <div class="action-card">
            <h3><i class="fas fa-users"></i> Students & Reports</h3>
            <p>View all students, filter by department, download PDF report cards or send them by email.</p>
            <a href="{{ url_for('ui_students') }}" class="btn">Open Students</a>
        </div>



        <div class="action-card">
            <h3><i class="fas fa-book"></i> Subjects & Codes</h3>
            <p>Maintain subject codes, names, year, semester and maximum marks for all departments.</p>
            <a href="{{ url_for('ui_subjects') }}" class="btn">Manage Subjects</a>
        </div>
    </div>
</div>

<footer>
    &copy; 2025 Sethupathy Government Arts College, Ramanathapuram. Powered by internal portal.
</footer>
</body>
</html>
"""

  # keep the full student form HTML from your last working version
NEW_SUBJECT_HTML = """
<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Add Subject | SGAC Student Portal</title>
    <link rel="stylesheet"
          href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap"
          rel="stylesheet">
    <style>
        :root {
            --primary-blue: #1f618d;
            --secondary-blue: #2e86c1;
            --accent-gold: #f1c40f;
            --bg-light: #f5f7fa;
            --text-dark: #333;
            --shadow-sm: 0 2px 4px rgba(0,0,0,0.05);
        }
        * { box-sizing:border-box; }
        body {
            margin:0;
            font-family:'Roboto',sans-serif;
            background:var(--bg-light);
            color:var(--text-dark);
        }
        a { text-decoration:none; color:inherit; }
        .top-bar {
            background:var(--primary-blue);
            color:#fff;
            font-size:0.9rem;
        }
        .top-bar-content {
            width:95%;
            max-width:800px;
            margin:0 auto;
            display:flex;
            justify-content:space-between;
            align-items:center;
            padding:5px 0;
        }
        .top-bar a { color:#fff; margin-right:15px; }
        .btn-sm-top {
            padding:4px 10px;
            border-radius:20px;
            background:#fff;
            color:var(--primary-blue) !important;
            font-size:0.8rem;
            font-weight:500;
        }
        .page-wrapper {
            width:95%;
            max-width:800px;
            margin:25px auto 40px;
        }
        .card {
            background:#fff;
            padding:20px;
            border-radius:8px;
            box-shadow:var(--shadow-sm);
            border-top:4px solid var(--primary-blue);
        }
        h1 {
            margin-top:0;
            font-size:1.4rem;
            color:var(--primary-blue);
            display:flex;
            align-items:center;
            gap:8px;
        }
        p { font-size:0.9rem; color:#555; }
        label {
            display:block;
            margin-bottom:4px;
            font-size:0.9rem;
        }
        input {
            width:100%;
            padding:7px;
            margin-bottom:10px;
            border-radius:4px;
            border:1px solid #ccc;
            font-size:0.9rem;
        }
        .row {
            display:flex;
            gap:10px;
            flex-wrap:wrap;
        }
        .row > div {
            flex:1;
            min-width:160px;
        }
        button {
            padding:8px 16px;
            border:none;
            border-radius:4px;
            background:var(--primary-blue);
            color:#fff;
            font-weight:500;
            cursor:pointer;
            font-size:0.9rem;
        }
        .link { color:var(--primary-blue); font-size:0.9rem; }
    </style>
</head>
<body>
<div class="top-bar">
    <div class="top-bar-content">
        <div>
            <a href="{{ url_for('ui_index') }}"><i class="fas fa-home"></i> Portal Home</a>
        </div>
        <div>
            <a href="{{ url_for('ui_students') }}" class="btn-sm-top">Students</a>
            <a href="{{ url_for('ui_new_student') }}" class="btn-sm-top">Add Student</a>
        </div>
    </div>
</div>

<div class="page-wrapper">
    <div class="card">
        <h1><i class="fas fa-book"></i> Add Subject</h1>
        <p>
            Define a subject with a unique code, year, semester and maximum marks.
            Use the same codes when entering student marks so reports stay consistent.
        </p>
        <form method="post">
            <label>Subject Name</label>
            <input name="name" required placeholder="e.g., Programming in C">

            <label>Subject Code</label>
            <input name="code" required placeholder="e.g., CS101 (must be unique)">

            <div class="row">
                <div>
                    <label>Year</label>
                    <input type="number" name="year" min="1" max="3" required>
                </div>
                <div>
                    <label>Semester</label>
                    <input type="number" name="semester" min="1" max="6" required>
                </div>
                <div>
                    <label>Max Marks</label>
                    <input type="number" name="max_marks" value="100" required>
                </div>
            </div>

            <button type="submit"><i class="fas fa-plus-circle"></i> Save Subject</button>
            <span style="margin-left:10px;">
                <a href="{{ url_for('ui_index') }}" class="link">Back to Home</a>
            </span>
        </form>
    </div>
</div>
</body>
</html>
"""

  # keep subject form
STUDENTS_LIST_HTML = """
<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Students | SGAC Student Portal</title>
    <link rel="stylesheet"
          href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap"
          rel="stylesheet">
    <style>
        :root {
            --primary-blue: #1f618d;
            --secondary-blue: #2e86c1;
            --accent-gold: #f1c40f;
            --bg-light: #f5f7fa;
            --text-dark: #333;
            --shadow-sm: 0 2px 4px rgba(0,0,0,0.05);
        }
        * { box-sizing:border-box; }
        body {
            margin:0;
            font-family:'Roboto',sans-serif;
            background:var(--bg-light);
            color:var(--text-dark);
        }
        a { text-decoration:none; color:inherit; }
        .top-bar {
            background:var(--primary-blue);
            color:#fff;
            font-size:0.9rem;
        }
        .top-bar-content {
            width:95%;
            max-width:1100px;
            margin:0 auto;
            display:flex;
            justify-content:space-between;
            align-items:center;
            padding:5px 0;
        }
        .top-bar a { color:#fff; margin-right:15px; }
        .btn-sm-top {
            padding:4px 10px;
            border-radius:20px;
            background:#fff;
            color:var(--primary-blue) !important;
            font-size:0.8rem;
            font-weight:500;
        }
        .page-wrapper {
            width:95%;
            max-width:1100px;
            margin:25px auto 40px;
        }
        .card {
            background:#fff;
            padding:20px;
            border-radius:8px;
            box-shadow:var(--shadow-sm);
            border-top:4px solid var(--primary-blue);
        }
        h1 {
            margin-top:0;
            font-size:1.4rem;
            color:var(--primary-blue);
            display:flex;
            align-items:center;
            gap:8px;
        }
        p { font-size:0.9rem; color:#555; }
        table {
            width:100%;
            border-collapse:collapse;
            font-size:0.85rem;
            margin-top:10px;
        }
        th, td {
            border:1px solid #eee;
            padding:6px;
            text-align:left;
        }
        th {
            background:#f0f3f9;
        }
        .badge {
            padding:2px 6px;
            border-radius:4px;
            font-size:0.75rem;
        }
        .badge-green { background:#e8f8f5; color:#16a085; }
        .badge-red { background:#fdecea; color:#c0392b; }
        button {
            padding:4px 8px;
            border:none;
            border-radius:4px;
            background:var(--primary-blue);
            color:#fff;
            font-size:0.75rem;
            cursor:pointer;
        }
        .filters {
            margin-bottom:10px;
            font-size:0.85rem;
            display:flex;
            gap:10px;
            flex-wrap:wrap;
        }
        .filters label { display:block; margin-bottom:2px; }
        .filters select,
        .filters input {
            padding:4px 6px;
            border-radius:4px;
            border:1px solid #ccc;
            font-size:0.85rem;
            min-width:120px;
        }
    </style>
</head>
<body>
<div class="top-bar">
    <div class="top-bar-content">
        <div>
            <a href="{{ url_for('ui_index') }}"><i class="fas fa-home"></i> Portal Home</a>
        </div>
        <div>
            <a href="{{ url_for('ui_new_student') }}" class="btn-sm-top">Add Student</a>
            <a href="{{ url_for('ui_subjects') }}" class="btn-sm-top">Subjects</a>
            <a href="{{ url_for('ui_email_settings') }}" class="btn-sm-top">Email Settings</a>
        </div>
    </div>
</div>

<div class="page-wrapper">
    <div class="card">
        <h1><i class="fas fa-users"></i> Students</h1>
        <p>
            View all students, filter by department/program/year, download PDF report cards or send them by email.
        </p>

        <div class="filters">
            <div>
                <label>Search</label>
                <input id="filter-text" placeholder="Name or ID">
            </div>
            <div>
                <label>Department</label>
                <select id="filter-dept">
                    <option value="">All</option>
                    {% for s in students|map(attribute='department')|unique %}
                      <option value="{{ s }}">{{ s }}</option>
                    {% endfor %}
                </select>
            </div>
            <div>
                <label>Program</label>
                <select id="filter-program">
                    <option value="">All</option>
                    <option value="UG">UG</option>
                    <option value="PG">PG</option>
                </select>
            </div>
            <div>
                <label>Year</label>
                <select id="filter-year">
                    <option value="">All</option>
                    <option value="1">I</option>
                    <option value="2">II</option>
                    <option value="3">III</option>
                </select>
            </div>
        </div>

        <table id="students-table">
            <thead>
                <tr>
                    <th>ID</th><th>Name</th><th>Dept</th>
                    <th>Program</th><th>Course</th>
                    <th>Year</th><th>Sem</th><th>Email</th>
                    <th>Attendance%</th><th>Marks%</th><th>PDF</th><th>Email</th>
                </tr>
            </thead>
            <tbody>
            {% for s in students %}
              <tr>
                <td>{{s.student_id}}</td>
                <td>{{s.name}}</td>
                <td>{{s.department}}</td>
                <td>{{s.program_type}}</td>
                <td>{{s.course_name}}</td>
                <td>{{s.year}}</td>
                <td>{{s.semester}}</td>
                <td>{{s.email}}</td>
                <td>
                    {% if s.attendance_percent|float >= 75 %}
                        <span class="badge badge-green">{{s.attendance_percent}}</span>
                    {% else %}
                        <span class="badge badge-red">{{s.attendance_percent}}</span>
                    {% endif %}
                </td>
                <td>{{s.marks_percent}}</td>
                <td>
                    <a href="{{ url_for('download_report', student_id=s.student_id) }}" target="_blank">
                        <i class="fas fa-file-pdf"></i> PDF
                    </a>
                </td>
                <td>
                    <form method="post" action="{{ url_for('ui_send_email', student_id=s.student_id) }}">
                        <button type="submit">
                            <i class="fas fa-paper-plane"></i> Send
                        </button>
                    </form>
                </td>
              </tr>
            {% endfor %}
            </tbody>
        </table>
    </div>
</div>

<script>
(function() {
    const textInput = document.getElementById('filter-text');
    const deptSelect = document.getElementById('filter-dept');
    const programSelect = document.getElementById('filter-program');
    const yearSelect = document.getElementById('filter-year');
    const table = document.getElementById('students-table');
    if (!table) return;
    const rows = Array.from(table.querySelectorAll('tbody tr'));

    function applyFilters() {
        const text = (textInput.value || "").toLowerCase();
        const dept = deptSelect.value;
        const program = programSelect.value;
        const year = yearSelect.value;

        rows.forEach(row => {
            const cells = row.children;
            const id = cells[0].textContent.toLowerCase();
            const name = cells[1].textContent.toLowerCase();
            const rdept = cells[2].textContent;
            const rprog = cells[3].textContent;
            const ryear = cells[5].textContent;

            let show = true;
            if (text && !(id.includes(text) || name.includes(text))) show = false;
            if (dept && rdept !== dept) show = false;
            if (program && rprog !== program) show = false;
            if (year && ryear !== year) show = false;

            row.style.display = show ? "" : "none";
        });
    }

    textInput.addEventListener('input', applyFilters);
    deptSelect.addEventListener('change', applyFilters);
    programSelect.addEventListener('change', applyFilters);
    yearSelect.addEventListener('change', applyFilters);
})();
</script>
</body>
</html>
"""

 # update below instead of "..."


SETTINGS_HTML = """
<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Email Settings | SGAC</title>
</head>
<body>
<h1>Email Settings (Gmail App Password)</h1>
<form method="post">
    <label>Gmail Address</label><br>
    <input name="smtp_email" type="email" value="{{ smtp_email or '' }}" required><br><br>
    <label>Gmail App Password</label><br>
    <input name="smtp_app_password" type="password" value="{{ smtp_app_password or '' }}" required><br><br>
    <button type="submit">Save Settings</button>
    <a href="{{ url_for('ui_index') }}">Back to Home</a>
</form>
</body>
</html>
"""

# Replace the last columns in STUDENTS_LIST_HTML to include Email button
STUDENTS_LIST_HTML = """
<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Students | SGAC Student Portal</title>
    <link rel="stylesheet"
          href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap"
          rel="stylesheet">
    <style>
        :root {
            --primary-blue: #1f618d;
            --secondary-blue: #2e86c1;
            --accent-gold: #f1c40f;
            --bg-light: #f5f7fa;
            --text-dark: #333;
            --shadow-sm: 0 2px 4px rgba(0,0,0,0.05);
        }
        * { box-sizing:border-box; }
        body {
            margin:0;
            font-family:'Roboto',sans-serif;
            background:var(--bg-light);
            color:var(--text-dark);
        }
        a { text-decoration:none; color:inherit; }
        .top-bar {
            background:var(--primary-blue);
            color:#fff;
            font-size:0.9rem;
        }
        .top-bar-content {
            width:95%;
            max-width:1100px;
            margin:0 auto;
            display:flex;
            justify-content:space-between;
            align-items:center;
            padding:5px 0;
        }
        .top-bar a { color:#fff; margin-right:15px; }
        .btn-sm-top {
            padding:4px 10px;
            border-radius:20px;
            background:#fff;
            color:var(--primary-blue) !important;
            font-size:0.8rem;
            font-weight:500;
        }
        .page-wrapper {
            width:95%;
            max-width:1100px;
            margin:25px auto 40px;
        }
        .card {
            background:#fff;
            padding:20px;
            border-radius:8px;
            box-shadow:var(--shadow-sm);
            border-top:4px solid var(--primary-blue);
        }
        h1 {
            margin-top:0;
            font-size:1.4rem;
            color:var(--primary-blue);
            display:flex;
            align-items:center;
            gap:8px;
        }
        p { font-size:0.9rem; color:#555; }
        table {
            width:100%;
            border-collapse:collapse;
            font-size:0.85rem;
            margin-top:10px;
        }
        th, td {
            border:1px solid #eee;
            padding:6px;
            text-align:left;
        }
        th {
            background:#f0f3f9;
        }
        .badge {
            padding:2px 6px;
            border-radius:4px;
            font-size:0.75rem;
        }
        .badge-green { background:#e8f8f5; color:#16a085; }
        .badge-red { background:#fdecea; color:#c0392b; }
        button {
            padding:4px 8px;
            border:none;
            border-radius:4px;
            background:var(--primary-blue);
            color:#fff;
            font-size:0.75rem;
            cursor:pointer;
        }
        .filters {
            margin-bottom:10px;
            font-size:0.85rem;
            display:flex;
            gap:10px;
            flex-wrap:wrap;
        }
        .filters label { display:block; margin-bottom:2px; }
        .filters select,
        .filters input {
            padding:4px 6px;
            border-radius:4px;
            border:1px solid #ccc;
            font-size:0.85rem;
            min-width:120px;
        }
    </style>
</head>
<body>
<div class="top-bar">
    <div class="top-bar-content">
        <div>
            <a href="{{ url_for('ui_index') }}"><i class="fas fa-home"></i> Portal Home</a>
        </div>
        <div>
            <a href="{{ url_for('ui_new_student') }}" class="btn-sm-top">Add Student</a>
            <a href="{{ url_for('ui_subjects') }}" class="btn-sm-top">Subjects</a>
            <a href="{{ url_for('ui_email_settings') }}" class="btn-sm-top">Email Settings</a>
        </div>
    </div>
</div>

<div class="page-wrapper">
    <div class="card">
        <h1><i class="fas fa-users"></i> Students</h1>
        <p>
            View all students, filter by department/program/year, download PDF report cards or send them by email.
        </p>

        <div class="filters">
            <div>
                <label>Search</label>
                <input id="filter-text" placeholder="Name or ID">
            </div>
            <div>
                <label>Department</label>
                <select id="filter-dept">
                    <option value="">All</option>
                    {% for s in students|map(attribute='department')|unique %}
                      <option value="{{ s }}">{{ s }}</option>
                    {% endfor %}
                </select>
            </div>
            <div>
                <label>Program</label>
                <select id="filter-program">
                    <option value="">All</option>
                    <option value="UG">UG</option>
                    <option value="PG">PG</option>
                </select>
            </div>
            <div>
                <label>Year</label>
                <select id="filter-year">
                    <option value="">All</option>
                    <option value="1">I</option>
                    <option value="2">II</option>
                    <option value="3">III</option>
                </select>
            </div>
        </div>

        <table id="students-table">
            <thead>
                <tr>
                    <th>ID</th><th>Name</th><th>Dept</th>
                    <th>Program</th><th>Course</th>
                    <th>Year</th><th>Sem</th><th>Email</th>
                    <th>Attendance%</th><th>Marks%</th><th>PDF</th><th>Email</th>
                </tr>
            </thead>
            <tbody>
            {% for s in students %}
              <tr>
                <td>{{s.student_id}}</td>
                <td>{{s.name}}</td>
                <td>{{s.department}}</td>
                <td>{{s.program_type}}</td>
                <td>{{s.course_name}}</td>
                <td>{{s.year}}</td>
                <td>{{s.semester}}</td>
                <td>{{s.email}}</td>
                <td>
                    {% if s.attendance_percent|float >= 75 %}
                        <span class="badge badge-green">{{s.attendance_percent}}</span>
                    {% else %}
                        <span class="badge badge-red">{{s.attendance_percent}}</span>
                    {% endif %}
                </td>
                <td>{{s.marks_percent}}</td>
                <td>
                    <a href="{{ url_for('download_report', student_id=s.student_id) }}" target="_blank">
                        <i class="fas fa-file-pdf"></i> PDF
                    </a>
                </td>
                <td>
                    <form method="post" action="{{ url_for('ui_send_email', student_id=s.student_id) }}">
                        <button type="submit">
                            <i class="fas fa-paper-plane"></i> Send
                        </button>
                    </form>
                </td>
              </tr>
            {% endfor %}
            </tbody>
        </table>
    </div>
</div>

<script>
(function() {
    const textInput = document.getElementById('filter-text');
    const deptSelect = document.getElementById('filter-dept');
    const programSelect = document.getElementById('filter-program');
    const yearSelect = document.getElementById('filter-year');
    const table = document.getElementById('students-table');
    if (!table) return;
    const rows = Array.from(table.querySelectorAll('tbody tr'));

    function applyFilters() {
        const text = (textInput.value || "").toLowerCase();
        const dept = deptSelect.value;
        const program = programSelect.value;
        const year = yearSelect.value;

        rows.forEach(row => {
            const cells = row.children;
            const id = cells[0].textContent.toLowerCase();
            const name = cells[1].textContent.toLowerCase();
            const rdept = cells[2].textContent;
            const rprog = cells[3].textContent;
            const ryear = cells[5].textContent;

            let show = true;
            if (text && !(id.includes(text) || name.includes(text))) show = false;
            if (dept && rdept !== dept) show = false;
            if (program && rprog !== program) show = false;
            if (year && ryear !== year) show = false;

            row.style.display = show ? "" : "none";
        });
    }

    textInput.addEventListener('input', applyFilters);
    deptSelect.addEventListener('change', applyFilters);
    programSelect.addEventListener('change', applyFilters);
    yearSelect.addEventListener('change', applyFilters);
})();
</script>
</body>
</html>
"""
NEW_STUDENT_HTML = """
<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Add / Update Student | SGAC</title>
    <link rel="stylesheet"
          href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap"
          rel="stylesheet">
    <style>
        :root {
            --primary-blue: #1f618d;
            --secondary-blue: #2e86c1;
            --accent-gold: #f1c40f;
            --bg-light: #f5f7fa;
            --text-dark: #333;
            --shadow-sm: 0 2px 4px rgba(0,0,0,0.05);
        }
        * { box-sizing:border-box; }
        body {
            margin:0;
            font-family:'Roboto',sans-serif;
            background:var(--bg-light);
            color:var(--text-dark);
        }
        a { text-decoration:none; color:inherit; }
        .top-bar {
            background:var(--primary-blue);
            color:#fff;
            font-size:0.9rem;
        }
        .top-bar-content {
            width:95%;
            max-width:1100px;
            margin:0 auto;
            display:flex;
            justify-content:space-between;
            align-items:center;
            padding:5px 0;
        }
        .top-bar a { color:#fff; margin-right:15px; }
        .btn-sm-top {
            padding:4px 10px;
            border-radius:20px;
            background:#fff;
            color:var(--primary-blue) !important;
            font-size:0.8rem;
            font-weight:500;
        }
        header {
            background:#fff;
            box-shadow:0 2px 4px rgba(0,0,0,0.05);
        }
        .header-container {
            width:95%;
            max-width:1100px;
            margin:0 auto;
            display:flex;
            justify-content:space-between;
            align-items:center;
            padding:10px 0;
        }
        .logo-section {
            display:flex;
            align-items:center;
            gap:10px;
        }
        .govt-logo, .college-logo {
            width:50px;
            height:50px;
            object-fit:contain;
        }
        .college-title h1 {
            margin:0;
            font-size:1.3rem;
            color:var(--primary-blue);
        }
        .college-title p {
            margin:2px 0;
            font-size:0.85rem;
            color:var(--secondary-blue);
        }
        .college-title span {
            font-size:0.75rem;
            color:#555;
        }
        .page-wrapper {
            width:95%;
            max-width:1100px;
            margin:20px auto 30px;
        }
        .card {
            background:#fff;
            padding:20px;
            border-radius:8px;
            box-shadow:var(--shadow-sm);
            border-top:4px solid var(--primary-blue);
        }
        h1 {
            margin-top:0;
            font-size:1.4rem;
            display:flex;
            align-items:center;
            gap:8px;
            color:var(--primary-blue);
        }
        label { display:block; margin-bottom:5px; font-size:0.9rem; }
        input, select {
            width:100%;
            padding:7px;
            margin-bottom:10px;
            border-radius:4px;
            border:1px solid #ccc;
            font-size:0.9rem;
        }
        button {
            padding:8px 16px;
            border:none;
            border-radius:4px;
            background:var(--primary-blue);
            color:#fff;
            font-weight:500;
            cursor:pointer;
            font-size:0.9rem;
        }
        .link {
            color:var(--primary-blue);
            font-size:0.9rem;
        }
        .row {
            display:flex;
            gap:10px;
            flex-wrap:wrap;
        }
        .row > div {
            flex:1;
            min-width:200px;
        }
        table {
            width:100%;
            border-collapse:collapse;
            font-size:0.85rem;
            margin-bottom:10px;
        }
        th, td {
            border:1px solid #eee;
            padding:5px;
            text-align:left;
        }
        th {
            background:#f0f3f9;
        }
        .btn-sm {
            padding:5px 10px;
            font-size:0.8rem;
            border-radius:4px;
        }
        .btn-secondary {
            background:#7f8c8d;
            color:#fff;
        }
        .section-title {
            font-size:1.05rem;
            margin-top:15px;
            color:var(--primary-blue);
        }
        .breadcrumb {
            font-size:0.85rem;
            margin-bottom:10px;
            color:#666;
        }
        .breadcrumb a { color:var(--primary-blue); }
        @media (max-width:768px) {
            .header-container { flex-direction:column; align-items:flex-start; gap:10px; }
        }
    </style>
</head>
<body>
<div class="top-bar">
    <div class="top-bar-content">
        <div>
            <a href="tel:+914567220268"><i class="fas fa-phone-alt"></i> +91-4567-220268</a>
            <a href="mailto:principal@sgacrmd.edu.in"><i class="fas fa-envelope"></i> principal@sgacrmd.edu.in</a>
        </div>
        <div>
            <a href="{{ url_for('ui_students') }}" class="btn-sm-top">Students List</a>
            <a href="{{ url_for('ui_index') }}" class="btn-sm-top">Portal Home</a>
        </div>
    </div>
</div>

<header>
    <div class="header-container">
        <div class="logo-section">
            <img src="https://sgacrmd.edu.in/assets/tn_logo.png" alt="TN Govt Logo" class="govt-logo">
            <div class="college-title">
                <h1>Sethupathy Government Arts College</h1>
                <p>Student Performance & Attendance Portal</p>
                <span>Ramanathapuram - 623 502, Tamil Nadu, India</span>
            </div>
            <img src="https://sgacrmd.edu.in/assets/logoclg.png" alt="College Logo" class="college-logo">
        </div>
    </div>
</header>

<div class="page-wrapper">
    <div class="breadcrumb">
        <a href="{{ url_for('ui_index') }}">Home</a> &raquo; Add / Update Student
    </div>

    <div class="card">
        <h1><i class="fas fa-user-plus"></i> Add / Update Student</h1>
        <p style="font-size:0.9rem;color:#555;margin-bottom:15px;">
            Enter basic details, academic program, marks and monthly attendance. This information will be used to
            calculate overall performance and generate the PDF report card.
        </p>

        <form id="student-form" method="post">
            <div class="row">
                <div>
                    <label>Student ID</label>
                    <input name="student_id" required placeholder="e.g., 22CS001">
                </div>
                <div>
                    <label>Student Name</label>
                    <input name="name" required placeholder="Full Name">
                </div>
            </div>

            <div class="row">
                <div>
                    <label>Email</label>
                    <input name="email" type="email" placeholder="student@example.com">
                </div>
                <div>
                    <label>Department</label>
                    <select name="dept_name">
                        {% for code, name in departments %}
                          <option value="{{name}}">{{name}}</option>
                        {% endfor %}
                    </select>
                </div>
            </div>

            <div class="row">
                <div>
                    <label>Program Type</label>
                    <select name="program_type">
                        <option value="UG">UG</option>
                        <option value="PG">PG</option>
                    </select>
                </div>
                <div>
                    <label>Duration</label>
                    <select name="duration_years">
                        <option value="3">3 Years</option>
                        <option value="2">2 Years</option>
                    </select>
                </div>
                <div>
                    <label>Course Name</label>
                    <input name="course_name" placeholder="B.Sc. Computer Science">
                </div>
            </div>

            <div class="row">
                <div>
                    <label>Year</label>
                    <input type="number" name="year" min="1" max="3" required>
                </div>
                <div>
                    <label>Semester</label>
                    <select name="semester">
                        {% for s in semesters %}
                          <option value="{{s}}">{{s}}</option>
                        {% endfor %}
                    </select>
                </div>
            </div>

            <h2 class="section-title"><i class="fas fa-book"></i> Subject Marks</h2>
            <p style="font-size:0.85rem;color:#555;">
                Add subject code and marks for each subject. Use the "+" button to add more rows.
            </p>
            <table id="marks-table">
                <thead>
                    <tr>
                        <th>Subject Code</th>
                        <th>Marks (out of 100)</th>
                        <th style="width:40px;"></th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><input name="marks_code" placeholder="CS101"></td>
                        <td><input name="marks_value" type="number" min="0" max="100" placeholder="0-100"></td>
                        <td><button type="button" class="btn-sm btn-secondary" onclick="removeRow(this)">X</button></td>
                    </tr>
                </tbody>
            </table>
            <button type="button" class="btn-sm btn-secondary" onclick="addMarksRow()">
                <i class="fas fa-plus"></i> Add Subject Row
            </button>

            <h2 class="section-title"><i class="fas fa-calendar-check"></i> Attendance</h2>
            <p style="font-size:0.85rem;color:#555;">
                Select month and enter classes held and attended for that month.
            </p>
            <table id="attendance-table">
                <thead>
                    <tr>
                        <th>Month</th>
                        <th>Classes Held</th>
                        <th>Classes Attended</th>
                        <th style="width:40px;"></th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>
                            <select name="att_month">
                                {% for m in months %}
                                  <option value="{{m}}">{{m}}</option>
                                {% endfor %}
                            </select>
                        </td>
                        <td><input name="att_held" type="number" min="0" placeholder="e.g., 20"></td>
                        <td><input name="att_attended" type="number" min="0" placeholder="e.g., 18"></td>
                        <td><button type="button" class="btn-sm btn-secondary" onclick="removeRow(this)">X</button></td>
                    </tr>
                </tbody>
            </table>
            <button type="button" class="btn-sm btn-secondary" onclick="addAttendanceRow()">
                <i class="fas fa-plus"></i> Add Month Row
            </button>

            <input type="hidden" name="marks_json" id="marks_json">
            <input type="hidden" name="attendance_json" id="attendance_json">

            <div style="margin-top:15px;display:flex;align-items:center;gap:12px;">
                <button type="submit"><i class="fas fa-save"></i> Save Student</button>
                <a href="{{ url_for('ui_students') }}" class="link">Go to Students List</a>
            </div>
        </form>
    </div>
</div>

<script>
function addMarksRow() {
    const tbody = document.querySelector('#marks-table tbody');
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td><input name="marks_code" placeholder="CS101"></td>
        <td><input name="marks_value" type="number" min="0" max="100" placeholder="0-100"></td>
        <td><button type="button" class="btn-sm btn-secondary" onclick="removeRow(this)">X</button></td>
    `;
    tbody.appendChild(tr);
}

function addAttendanceRow() {
    const tbody = document.querySelector('#attendance-table tbody');
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td>
            <select name="att_month">
                {% for m in months %}
                  <option value="{{m}}">{{m}}</option>
                {% endfor %}
            </select>
        </td>
        <td><input name="att_held" type="number" min="0" placeholder="e.g., 20"></td>
        <td><input name="att_attended" type="number" min="0" placeholder="e.g., 18"></td>
        <td><button type="button" class="btn-sm btn-secondary" onclick="removeRow(this)">X</button></td>
    `;
    tbody.appendChild(tr);
}

function removeRow(btn) {
    const tr = btn.closest('tr');
    tr.remove();
}

document.getElementById('student-form').addEventListener('submit', function () {
    const marksData = {};
    const codes = document.getElementsByName('marks_code');
    const values = document.getElementsByName('marks_value');
    for (let i = 0; i < codes.length; i++) {
        const code = codes[i].value.trim();
        const val = values[i].value.trim();
        if (code && val) {
            marksData[code] = parseFloat(val);
        }
    }

    const attendanceData = {};
    const months = document.getElementsByName('att_month');
    const helds = document.getElementsByName('att_held');
    const attendeds = document.getElementsByName('att_attended');
    for (let i = 0; i < months.length; i++) {
        const m = months[i].value;
        const h = helds[i].value.trim();
        const a = attendeds[i].value.trim();
        if (m && h) {
            attendanceData[m] = {
                held: parseInt(h || "0"),
                attended: parseInt(a || "0")
            };
        }
    }

    document.getElementById('marks_json').value = JSON.stringify(marksData);
    document.getElementById('attendance_json').value = JSON.stringify(attendanceData);
});
</script>
</body>
</html>
"""



@app.get("/")
def ui_index():
    tn_logo = "/" + TN_LOGO_PATH if os.path.exists(TN_LOGO_PATH) else "https://via.placeholder.com/60"
    college_logo = "/" + COLLEGE_LOGO_PATH if os.path.exists(COLLEGE_LOGO_PATH) else "https://via.placeholder.com/60"
    return render_template_string(
        INDEX_HTML,
        tn_logo=tn_logo,
        college_logo=college_logo,
    )


@app.route("/ui/new-student", methods=["GET", "POST"])
def ui_new_student():
    import json

    if request.method == "GET":
        departments = system.get_all_departments()
        return render_template_string(
            NEW_STUDENT_HTML,
            departments=departments,
            months=MONTHS,
            semesters=SEMESTERS,
        )

    form = request.form
    student_id = form.get("student_id", "").strip()
    name = form.get("name", "").strip()
    email = form.get("email", "").strip()
    dept_name = form.get("dept_name", "").strip()
    program_type = form.get("program_type", "").strip()
    duration_years = int(form.get("duration_years", 0))
    course_name = form.get("course_name", "").strip()
    year = int(form.get("year", 0))
    semester = int(form.get("semester", 0))
    marks_json = form.get("marks_json", "") or "{}"
    attendance_json = form.get("attendance_json", "") or "{}"

    try:
        marks_data = json.loads(marks_json)
        attendance_data = json.loads(attendance_json)
    except json.JSONDecodeError:
        return "Invalid data in marks or attendance", 400

    ok, msg = system.add_or_update_student(
        student_id,
        name,
        dept_name,
        year,
        semester,
        email,
        program_type,
        duration_years,
        course_name,
        marks_data,
        attendance_data,
    )
    status = 200 if ok else 400
    return f"<p>{msg}</p><p><a href='{url_for('ui_new_student')}'>Back to form</a> | <a href='{url_for('ui_index')}'>Home</a></p>", status


@app.route("/ui/subjects", methods=["GET", "POST"])
def ui_subjects():
    if request.method == "GET":
        return render_template_string(NEW_SUBJECT_HTML)

    name = request.form.get("name", "").strip()
    code = request.form.get("code", "").strip().upper()
    year = int(request.form.get("year", 0))
    semester = int(request.form.get("semester", 0))
    max_marks = int(request.form.get("max_marks", 100))
    ok, msg = system.add_subject(name, code, year, semester, max_marks)
    status = 200 if ok else 400
    return f"<p>{msg}</p><p><a href='{url_for('ui_index')}'>Back to Home</a></p>", status


@app.get("/ui/students")
def ui_students():
    data = system.get_all_students_summary()
    students = [
        {
            "student_id": row[0],
            "name": row[1],
            "department": row[2],
            "year": row[3],
            "semester": row[4],
            "email": row[5],
            "program_type": row[6],
            "duration_years": row[7],
            "course_name": row[8],
            "attendance_percent": row[9],
            "marks_percent": row[10],
        }
        for row in data
    ]
    return render_template_string(STUDENTS_LIST_HTML, students=students)


@app.route("/ui/settings/email", methods=["GET", "POST"])
def ui_email_settings():
    if request.method == "GET":
        email, app_pwd = system.get_email_settings()
        return render_template_string(
            SETTINGS_HTML,
            smtp_email=email,
            smtp_app_password=app_pwd,
        )

    smtp_email = request.form.get("smtp_email", "").strip()
    smtp_app_password = request.form.get("smtp_app_password", "").strip()
    system.save_email_settings(smtp_email, smtp_app_password)
    return (
        "<p>Settings saved.</p>"
        f"<p><a href='{url_for('ui_email_settings')}'>Back to Settings</a> | "
        f"<a href='{url_for('ui_index')}'>Home</a></p>"
    )


@app.post("/ui/students/<student_id>/send-email")
def ui_send_email(student_id):
    ok, msg = system.email_report_to_student(student_id)
    status = 200 if ok else 400
    return (
        f"<p>{msg}</p>"
        f"<p><a href='{url_for('ui_students')}'>Back to Students</a> | "
        f"<a href='{url_for('ui_index')}'>Home</a></p>",
        status,
    )


@app.get("/students")
def api_list_students():
    data = system.get_all_students_summary()
    students = [
        {
            "student_id": row[0],
            "name": row[1],
            "department": row[2],
            "year": row[3],
            "semester": row[4],
            "email": row[5],
            "program_type": row[6],
            "duration_years": row[7],
            "course_name": row[8],
            "attendance_percent": row[9],
            "marks_percent": row[10],
        }
        for row in data
    ]
    return jsonify(students)


@app.get("/students/<student_id>")
def api_get_student(student_id):
    student, marks, attendance = system.get_student_details_and_data(student_id)
    if not student:
        abort(404, description="Student not found")
    return jsonify({"student": student, "marks": marks, "attendance": attendance})


@app.post("/students")
def api_create_or_update_student():
    payload = request.get_json(force=True)
    student_id = payload.get("student_id", "").strip()
    name = payload.get("name", "").strip()
    dept_name = payload.get("dept_name", "").strip()
    year = int(payload.get("year", 0))
    semester = int(payload.get("semester", 0))
    email = payload.get("email", "").strip()
    program_type = payload.get("program_type", "").strip()
    duration_years = int(payload.get("duration_years", 0))
    course_name = payload.get("course_name", "").strip()
    marks_data = payload.get("marks_data", {}) or {}
    attendance_data = payload.get("attendance_data", {}) or {}

    ok, msg = system.add_or_update_student(
        student_id,
        name,
        dept_name,
        year,
        semester,
        email,
        program_type,
        duration_years,
        course_name,
        marks_data,
        attendance_data,
    )
    status = 200 if ok else 400
    return jsonify({"success": ok, "message": msg}), status


@app.get("/subjects")
def api_get_subjects():
    subs = system.get_all_subjects()
    subjects = [
        {
            "subject_code": row[0],
            "subject_name": row[1],
            "year": row[2],
            "semester": row[3],
            "max_marks": row[4],
        }
        for row in subs
    ]
    return jsonify(subjects)


@app.post("/subjects")
def api_add_subject():
    payload = request.get_json(force=True)
    name = payload.get("name", "").strip()
    code = payload.get("code", "").strip().upper()
    year = int(payload.get("year", 0))
    semester = int(payload.get("semester", 0))
    max_marks = int(payload.get("max_marks", 100))
    ok, msg = system.add_subject(name, code, year, semester, max_marks)
    status = 200 if ok else 400
    return jsonify({"success": ok, "message": msg}), status


@app.get("/reports/<student_id>.pdf")
def download_report(student_id):
    if not REPORTLAB_INSTALLED:
        abort(500, description="reportlab not installed on server")
    try:
        pdf_path = system.generate_student_report_pdf(student_id)
    except ValueError as e:
        abort(404, description=str(e))
    except Exception as e:
        abort(500, description=f"PDF error: {e}")
    if not os.path.exists(pdf_path):
        abort(404, description="PDF not found")
    return send_file(pdf_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
