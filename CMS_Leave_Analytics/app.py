from flask import Flask, render_template, request, jsonify, send_file
import sqlite3
from pathlib import Path
from datetime import datetime
import pandas as pd
import io
import re

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "cms_leave_analytics.db"
UPLOAD_DIR = APP_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

# -----------------------------
# Database
# -----------------------------
def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = db()
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS employees (
        emp_code TEXT PRIMARY KEY,
        emp_name TEXT,
        date_of_joining TEXT,
        branch TEXT,
        department TEXT,
        sub_department TEXT,
        designation TEXT,
        functional_area TEXT
    );

    CREATE TABLE IF NOT EXISTS leaves (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        emp_code TEXT,
        emp_name TEXT,
        branch TEXT,
        designation TEXT,
        leave_type TEXT,
        token_no TEXT,
        requested_date TEXT,
        leave_from TEXT,
        leave_to TEXT,
        requested_days REAL DEFAULT 0,
        approved_days REAL DEFAULT 0,
        lop_days REAL DEFAULT 0,
        advance_days REAL DEFAULT 0,
        reason TEXT,
        leave_reason TEXT,
        leave_sub_reason TEXT,
        status TEXT,
        action_by TEXT,
        action_date TEXT,
        month TEXT,
        FOREIGN KEY(emp_code) REFERENCES employees(emp_code)
    );

    CREATE TABLE IF NOT EXISTS uploads (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT,
        uploaded_at TEXT,
        rows_imported INTEGER
    );
    """)
    conn.commit()
    conn.close()

init_db()

# -----------------------------
# Helpers
# -----------------------------
COL_MAP = {
    "E. Code": "emp_code",
    "Emp Name": "emp_name",
    "Date Of Joining": "date_of_joining",
    "Branch": "branch",
    "Department": "department",
    "Sub Department": "sub_department",
    "Designation": "designation",
    "Functional Area": "functional_area",
    "Leave Type": "leave_type",
    "Token_No": "token_no",
    "Requested Date": "requested_date",
    "Leave From": "leave_from",
    "Leave To": "leave_to",
    "Requested Days": "requested_days",
    "Approved Days": "approved_days",
    "LOP Days": "lop_days",
    "Advance Days": "advance_days",
    "Reason": "reason",
    "Leave Reason": "leave_reason",
    "Leave Sub Reason": "leave_sub_reason",
    "Status": "status",
    "Action By": "action_by",
    "Action Date": "action_date",
}

def clean_value(v):
    if pd.isna(v):
        return ""
    return str(v).strip()

def normalize_date(v):
    if not v:
        return ""
    d = pd.to_datetime(v, errors="coerce", dayfirst=True)
    if pd.isna(d):
        return clean_value(v)
    return d.strftime("%Y-%m-%d")

def parse_report(file_path):
    # HRMS exports can be HTML tables saved with an .xls extension.
    try:
        tables = pd.read_html(file_path)
        if tables:
            df = tables[0]
        else:
            raise ValueError("No table found.")
    except Exception:
        df = pd.read_excel(file_path)

    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in COL_MAP if c not in df.columns]
    if missing:
        raise ValueError("Required columns are missing: " + ", ".join(missing))

    df = df[list(COL_MAP.keys())].rename(columns=COL_MAP)

    for col in ["emp_code", "emp_name", "branch", "designation", "leave_type",
                "token_no", "reason", "leave_reason", "leave_sub_reason",
                "status", "action_by", "date_of_joining"]:
        df[col] = df[col].map(clean_value)

    for col in ["requested_days", "approved_days", "lop_days", "advance_days"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    for col in ["requested_date", "leave_from", "leave_to", "action_date"]:
        df[col] = df[col].map(normalize_date)

    # Month is based on Leave From, which is the useful month for monthly reporting.
    parsed = pd.to_datetime(df["leave_from"], errors="coerce")
    df["month"] = parsed.dt.strftime("%Y-%m").fillna("Unknown")

    return df

def import_dataframe(df, filename):
    conn = db()
    cur = conn.cursor()

    # Replace the current report with the latest imported report.
    cur.execute("DELETE FROM leaves")
    cur.execute("DELETE FROM employees")

    emp_cols = ["emp_code", "emp_name", "date_of_joining", "branch",
                "department", "sub_department", "designation", "functional_area"]

    employees = df[emp_cols].drop_duplicates(subset=["emp_code"])
    for _, r in employees.iterrows():
        cur.execute("""
            INSERT OR REPLACE INTO employees
            (emp_code, emp_name, date_of_joining, branch, department,
             sub_department, designation, functional_area)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, tuple(clean_value(r[c]) for c in emp_cols))

    leave_cols = [
        "emp_code","emp_name","branch","designation","leave_type","token_no",
        "requested_date","leave_from","leave_to","requested_days","approved_days",
        "lop_days","advance_days","reason","leave_reason","leave_sub_reason",
        "status","action_by","action_date","month"
    ]
    for _, r in df.iterrows():
        values = []
        for c in leave_cols:
            if c in ["requested_days","approved_days","lop_days","advance_days"]:
                values.append(float(r[c] or 0))
            else:
                values.append(clean_value(r[c]))
        cur.execute(f"""
            INSERT INTO leaves ({",".join(leave_cols)})
            VALUES ({",".join(["?"] * len(leave_cols))})
        """, values)

    cur.execute(
        "INSERT INTO uploads(filename, uploaded_at, rows_imported) VALUES (?, ?, ?)",
        (filename, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), len(df))
    )
    conn.commit()
    conn.close()

def filters_from_request():
    return {
        "month": request.args.get("month", "").strip(),
        "leave_type": request.args.get("leave_type", "").strip(),
        "branch": request.args.get("branch", "").strip(),
        "designation": request.args.get("designation", "").strip(),
        "status": request.args.get("status", "").strip(),
        "search": request.args.get("search", "").strip(),
        "threshold": request.args.get("threshold", "2").strip(),
        "mode": request.args.get("mode", "monthly_total").strip(),
    }

def build_where(f):
    conditions = []
    params = []

    if f["month"]:
        conditions.append("month = ?")
        params.append(f["month"])
    if f["leave_type"]:
        conditions.append("leave_type = ?")
        params.append(f["leave_type"])
    if f["branch"]:
        conditions.append("branch = ?")
        params.append(f["branch"])
    if f["designation"]:
        conditions.append("designation = ?")
        params.append(f["designation"])
    if f["status"]:
        conditions.append("status = ?")
        params.append(f["status"])
    if f["search"]:
        conditions.append("(emp_name LIKE ? OR emp_code LIKE ?)")
        q = f"%{f['search']}%"
        params.extend([q, q])

    where = " WHERE " + " AND ".join(conditions) if conditions else ""
    return where, params

# -----------------------------
# Pages / API
# -----------------------------
@app.route("/")
def index():
    conn = db()
    months = [r[0] for r in conn.execute(
        "SELECT DISTINCT month FROM leaves WHERE month <> '' ORDER BY month DESC"
    ).fetchall()]
    leave_types = [r[0] for r in conn.execute(
        "SELECT DISTINCT leave_type FROM leaves WHERE leave_type <> '' ORDER BY leave_type"
    ).fetchall()]
    branches = [r[0] for r in conn.execute(
        "SELECT DISTINCT branch FROM leaves WHERE branch <> '' ORDER BY branch"
    ).fetchall()]
    designations = [r[0] for r in conn.execute(
        "SELECT DISTINCT designation FROM leaves WHERE designation <> '' ORDER BY designation"
    ).fetchall()]
    statuses = [r[0] for r in conn.execute(
        "SELECT DISTINCT status FROM leaves WHERE status <> '' ORDER BY status"
    ).fetchall()]
    last_upload = conn.execute(
        "SELECT * FROM uploads ORDER BY id DESC LIMIT 1"
    ).fetchone()
    conn.close()

    return render_template(
        "index.html",
        months=months,
        leave_types=leave_types,
        branches=branches,
        designations=designations,
        statuses=statuses,
        last_upload=last_upload
    )

@app.post("/upload")
def upload():
    file = request.files.get("report")
    if not file or not file.filename:
        return jsonify(ok=False, error="Please select an HRMS report."), 400

    safe_name = re.sub(r"[^A-Za-z0-9._-]", "_", file.filename)
    path = UPLOAD_DIR / safe_name
    file.save(path)

    try:
        df = parse_report(path)
        import_dataframe(df, file.filename)
        return jsonify(ok=True, rows=len(df), message=f"Imported {len(df)} leave records successfully.")
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 400

@app.get("/api/dashboard")
def dashboard():
    f = filters_from_request()
    where, params = build_where(f)
    threshold = max(0.0, float(f["threshold"] or 2))

    conn = db()

    totals = conn.execute(f"""
        SELECT
            COUNT(*) AS records,
            COUNT(DISTINCT emp_code) AS employees,
            COALESCE(SUM(approved_days),0) AS approved_days,
            COALESCE(SUM(lop_days),0) AS lop_days
        FROM leaves {where}
    """, params).fetchone()

    # Monthly employee totals. This is the main "more than 2 days in a month" report.
    monthly_sql = f"""
        SELECT emp_code, emp_name, branch, designation, month,
               ROUND(SUM(approved_days),2) AS total_days,
               COUNT(*) AS leave_records
        FROM leaves {where}
        GROUP BY emp_code, emp_name, branch, designation, month
        HAVING SUM(approved_days) >= ?
        ORDER BY total_days DESC, emp_name
    """
    monthly = conn.execute(monthly_sql, params + [threshold]).fetchall()

    # Individual leave applications longer than the selected threshold.
    individual_sql = f"""
        SELECT emp_code, emp_name, branch, designation, leave_type,
               leave_from, leave_to, approved_days, status, reason
        FROM leaves {where}
        AND approved_days >= ?
        ORDER BY approved_days DESC, leave_from
    """ if where else """
        SELECT emp_code, emp_name, branch, designation, leave_type,
               leave_from, leave_to, approved_days, status, reason
        FROM leaves
        WHERE approved_days >= ?
        ORDER BY approved_days DESC, leave_from
    """
    individual = conn.execute(individual_sql, params + [threshold]).fetchall()

    by_type = conn.execute(f"""
        SELECT leave_type, ROUND(SUM(approved_days),2) AS days
        FROM leaves {where}
        GROUP BY leave_type
        ORDER BY days DESC
    """, params).fetchall()

    by_month = conn.execute(f"""
        SELECT month, ROUND(SUM(approved_days),2) AS days
        FROM leaves {where}
        GROUP BY month
        ORDER BY month
    """, params).fetchall()

    top_employees = conn.execute(f"""
        SELECT emp_name, ROUND(SUM(approved_days),2) AS days
        FROM leaves {where}
        GROUP BY emp_code, emp_name
        ORDER BY days DESC
        LIMIT 10
    """, params).fetchall()

    conn.close()

    return jsonify({
        "totals": dict(totals),
        "monthly_over_2": [dict(r) for r in monthly],
        "individual_over_2": [dict(r) for r in individual],
        "by_type": [dict(r) for r in by_type],
        "by_month": [dict(r) for r in by_month],
        "top_employees": [dict(r) for r in top_employees],
        "threshold": threshold
    })

@app.get("/export/csv")
def export_csv():
    f = filters_from_request()
    where, params = build_where(f)
    threshold = max(0.0, float(f["threshold"] or 2))

    conn = db()
    if f["mode"] == "individual":
        sql = f"""
            SELECT emp_code AS "Employee Code", emp_name AS "Employee Name",
                   branch AS "Branch", designation AS "Designation",
                   leave_type AS "Leave Type", leave_from AS "Leave From",
                   leave_to AS "Leave To", approved_days AS "Approved Days",
                   status AS "Status", reason AS "Reason"
            FROM leaves {where}
            {'AND' if where else 'WHERE'} approved_days >= ?
            ORDER BY approved_days DESC
        """
        rows = conn.execute(sql, params + [threshold]).fetchall()
        title = "Long Leave Applications"
    else:
        sql = f"""
            SELECT emp_code AS "Employee Code", emp_name AS "Employee Name",
                   branch AS "Branch", designation AS "Designation",
                   month AS "Month", ROUND(SUM(approved_days),2) AS "Total Approved Days",
                   COUNT(*) AS "Leave Records"
            FROM leaves {where}
            GROUP BY emp_code, emp_name, branch, designation, month
            HAVING SUM(approved_days) >= ?
            ORDER BY "Total Approved Days" DESC
        """
        rows = conn.execute(sql, params + [threshold]).fetchall()
        title = "Employees With Leave Above Threshold Per Month"
    conn.close()

    df = pd.DataFrame([dict(r) for r in rows])
    out = io.StringIO()
    df.to_csv(out, index=False)
    out.seek(0)

    return send_file(
        io.BytesIO(out.getvalue().encode("utf-8-sig")),
        mimetype="text/csv",
        as_attachment=True,
        download_name=title.replace(" ", "_") + ".csv"
    )

@app.get("/health")
def health():
    return {"status": "ok"}

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
