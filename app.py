import os
import io
import datetime
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, UserMixin, login_user,
    login_required, logout_user, current_user
)
from werkzeug.security import generate_password_hash, check_password_hash
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import r2_score
import plotly.graph_objects as go
import json
import anthropic
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# -------------------------------------------------
# APP CONFIGURATION
# -------------------------------------------------

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "super-secret-key")

database_url = os.environ.get("DATABASE_URL", "sqlite:///local.db")
if database_url.startswith("postgres://"):
    database_url = database_url.replace("postgres://", "postgresql://")

app.config["SQLALCHEMY_DATABASE_URI"] = database_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

# -------------------------------------------------
# LOGIN MANAGER
# -------------------------------------------------

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"
login_manager.login_message = "Please log in first."

# -------------------------------------------------
# MODELS
# -------------------------------------------------

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(120), unique=True, nullable=False)
    password = db.Column(db.String(255), nullable=False)


class HospitalMetric(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, nullable=False)
    metric_name = db.Column(db.String(120), nullable=False)
    metric_value = db.Column(db.Float, nullable=False)


@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


# -------------------------------------------------
# AUTH ROUTES
# -------------------------------------------------

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = request.form["username"]
        password = generate_password_hash(request.form["password"])

        if User.query.filter_by(username=username).first():
            flash("Username already exists.")
            return redirect(url_for("register"))

        new_user = User(username=username, password=password)
        db.session.add(new_user)
        db.session.commit()
        flash("Account created. Please login.")
        return redirect(url_for("login"))

    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]
        user = User.query.filter_by(username=username).first()

        if user and check_password_hash(user.password, password):
            login_user(user)
            return redirect(url_for("dashboard"))
        else:
            flash("Invalid username or password.")

    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


# -------------------------------------------------
# HELPER: BUILD ANALYSIS SUMMARY FROM DB
# -------------------------------------------------

def build_analysis_summary():
    records = HospitalMetric.query.all()
    if not records:
        return None

    df = pd.DataFrame([{
        "date": r.date, "metric": r.metric_name, "value": r.metric_value
    } for r in records])

    total_cases   = int(df["value"].sum())
    metric_totals = df.groupby("metric")["value"].sum().sort_values(ascending=False)
    top_disease   = metric_totals.index[0]

    df["month"] = pd.to_datetime(df["date"]).dt.strftime("%B")
    peak_month = df.groupby("month")["value"].sum().idxmax()

    df_grouped = df.groupby(["date", "metric"])["value"].sum().reset_index()
    df_grouped["MonthIndex"] = df_grouped.groupby("metric").cumcount() + 1

    top_df = df_grouped[df_grouped["metric"] == top_disease].copy()
    X = top_df[["MonthIndex"]]
    y = top_df["value"]

    model = RandomForestRegressor(n_estimators=100, random_state=42)
    model.fit(X, y)
    predictions  = model.predict(X)
    accuracy     = round(r2_score(y, predictions) * 100, 1)

    case_distribution = []
    for metric, total in metric_totals.items():
        pct = round(total / metric_totals.sum() * 100, 1)
        case_distribution.append({"metric": metric, "total": int(total), "pct": pct})

    min_date   = df["date"].min()
    max_date   = df["date"].max()
    date_range = f"{pd.to_datetime(min_date).strftime('%b %Y')} – {pd.to_datetime(max_date).strftime('%b %Y')}"

    return {
        "total_cases":        total_cases,
        "top_disease":        top_disease,
        "peak_month":         peak_month,
        "accuracy":           accuracy,
        "case_distribution":  case_distribution,
        "date_range":         date_range,
    }


# -------------------------------------------------
# HELPER: CALL CLAUDE FOR ANALYSIS
# -------------------------------------------------

def get_ai_analysis(summary):
    dist_text = "\n".join(
        f"  - {d['metric']}: {d['total']:,} cases ({d['pct']}%)"
        for d in summary["case_distribution"]
    )

    prompt = f"""You are a senior medical data analyst reviewing hospital disease surveillance data for Kenya.

Here is the analysis summary:
- Date range: {summary['date_range']}
- Total reported cases: {summary['total_cases']:,}
- Highest disease burden: {summary['top_disease']}
- Peak reporting month: {summary['peak_month']}
- ML model accuracy: {summary['accuracy']}%

Disease distribution:
{dist_text}

Provide a detailed professional medical report with these exact sections:

1. Executive Summary
2. Key Findings
3. Disease Burden Analysis
4. Risk Alerts
5. Recommended Actions
6. Resource Allocation Suggestions
7. Forecast Outlook

Be specific, data-driven, and actionable. Write in full professional sentences."""

    client  = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt}]
    )
    return message.content[0].text


# -------------------------------------------------
# HELPER: BUILD WORD DOCUMENT
# -------------------------------------------------

def build_word_report(summary, ai_text):
    doc = Document()

    # ── Page margins ──
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.25)
        section.right_margin  = Inches(1.25)

    # ── Styles ──
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    # ── COVER ──────────────────────────────────────
    # Hospital name
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run("AMOS HOSPITAL ANALYTICS")
    run.bold      = True
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0x0D, 0x47, 0xA1)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = sub_p.add_run("Disease Surveillance & Analysis Report")
    run2.font.size = Pt(14)
    run2.font.color.rgb = RGBColor(0x37, 0x47, 0x5A)

    doc.add_paragraph()

    # Date line
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_p.add_run(f"Report Period: {summary['date_range']}").bold = True

    gen_p = doc.add_paragraph()
    gen_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    gen_p.add_run(f"Generated: {datetime.datetime.now().strftime('%B %d, %Y at %H:%M')}")

    doc.add_paragraph()
    doc.add_paragraph()

    # ── KPI SUMMARY TABLE ───────────────────────────
    h = doc.add_heading("Summary Statistics", level=1)
    h.runs[0].font.color.rgb = RGBColor(0x0D, 0x47, 0xA1)

    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"

    hdr_cells = table.rows[0].cells
    headers   = ["Total Cases", "Highest Disease", "Peak Month", "Model Accuracy"]
    values    = [
        f"{summary['total_cases']:,}",
        summary["top_disease"],
        summary["peak_month"],
        f"{summary['accuracy']}%"
    ]

    for i, (hdr, val) in enumerate(zip(headers, values)):
        cell = hdr_cells[i]
        # Header row
        hp = cell.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hr = hp.add_run(hdr)
        hr.bold = True
        hr.font.size = Pt(9)
        hr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        # Set blue background
        tc  = cell._tc
        tcp = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "1565C0")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:val"), "clear")
        tcp.append(shd)

        # Value row
        vrow = table.add_row()
        vc   = vrow.cells[i]
        vp   = vc.paragraphs[0]
        vp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        vr = vp.add_run(val)
        vr.bold      = True
        vr.font.size = Pt(13)

    doc.add_paragraph()

    # ── DISEASE DISTRIBUTION TABLE ──────────────────
    h2 = doc.add_heading("Disease Distribution", level=1)
    h2.runs[0].font.color.rgb = RGBColor(0x0D, 0x47, 0xA1)

    dtable = doc.add_table(rows=1, cols=3)
    dtable.style = "Table Grid"

    dhdrs = ["Rank", "Disease / Metric", "Cases", "Percentage"]
    dtable = doc.add_table(rows=1, cols=4)
    dtable.style = "Table Grid"

    for i, dh in enumerate(dhdrs):
        cell = dtable.rows[0].cells[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = cell.paragraphs[0].add_run(dh)
        r.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        tc  = cell._tc
        tcp = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "1565C0")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:val"), "clear")
        tcp.append(shd)

    for idx, item in enumerate(summary["case_distribution"], 1):
        row   = dtable.add_row().cells
        fill  = "EBF5FB" if idx % 2 == 0 else "FFFFFF"
        data  = [str(idx), item["metric"], f"{item['total']:,}", f"{item['pct']}%"]
        for ci, val in enumerate(data):
            row[ci].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row[ci].paragraphs[0].add_run(val).font.size = Pt(10)
            if fill == "EBF5FB":
                tc  = row[ci]._tc
                tcp = tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:fill"), fill)
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:val"), "clear")
                tcp.append(shd)

    doc.add_paragraph()
    doc.add_page_break()

    # ── AI ANALYSIS SECTIONS ────────────────────────
    h3 = doc.add_heading("AI-Powered Analysis & Recommendations", level=1)
    h3.runs[0].font.color.rgb = RGBColor(0x0D, 0x47, 0xA1)

    note_p = doc.add_paragraph()
    note_r = note_p.add_run("The following analysis was generated by Claude AI based on your uploaded data.")
    note_r.italic     = True
    note_r.font.size  = Pt(10)
    note_r.font.color.rgb = RGBColor(0x78, 0x90, 0x9C)

    doc.add_paragraph()

    # Parse AI text and write sections
    current_section = None
    for line in ai_text.split("\n"):
        line = line.strip()
        if not line:
            continue

        # Detect numbered headings like "1. Executive Summary" or "## Executive Summary"
        is_heading = (
            (line[0].isdigit() and ". " in line[:4]) or
            line.startswith("##") or
            line.startswith("**") and line.endswith("**")
        )

        if is_heading:
            clean = line.lstrip("#").lstrip("0123456789.").strip().strip("*").strip()
            h = doc.add_heading(clean, level=2)
            h.runs[0].font.color.rgb = RGBColor(0x15, 0x65, 0xC0)
        elif line.startswith("- ") or line.startswith("• "):
            p = doc.add_paragraph(style="List Bullet")
            p.add_run(line.lstrip("- ").lstrip("• ")).font.size = Pt(11)
        elif line.startswith("**") and "**" in line[2:]:
            p = doc.add_paragraph()
            r = p.add_run(line.strip("*"))
            r.bold      = True
            r.font.size = Pt(11)
        else:
            p = doc.add_paragraph()
            p.add_run(line).font.size = Pt(11)

    doc.add_paragraph()

    # ── FOOTER NOTE ─────────────────────────────────
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = footer_p.add_run(
        f"Confidential — Amos Hospital Analytics · Generated {datetime.datetime.now().strftime('%Y-%m-%d')}"
    )
    fr.font.size = Pt(9)
    fr.font.color.rgb = RGBColor(0x9E, 0x9E, 0x9E)
    fr.italic = True

    # Save to buffer
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# -------------------------------------------------
# GENERATE REPORT ROUTE
# -------------------------------------------------

@app.route("/generate-report")
@login_required
def generate_report():
    try:
        summary = build_analysis_summary()
        if not summary:
            flash("No data found. Please upload a CSV first.")
            return redirect(url_for("dashboard"))

        ai_text  = get_ai_analysis(summary)
        doc_buf  = build_word_report(summary, ai_text)
        filename = f"Amos_Hospital_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.docx"

        return send_file(
            doc_buf,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except anthropic.AuthenticationError:
        flash("Invalid Anthropic API key. Set ANTHROPIC_API_KEY and restart.")
        return redirect(url_for("dashboard"))
    except Exception as e:
        flash(f"Report error: {str(e)}")
        return redirect(url_for("dashboard"))


# -------------------------------------------------
# AI INSIGHTS (JSON — for the inline panel)
# -------------------------------------------------

@app.route("/ai-insights", methods=["POST"])
@login_required
def ai_insights():
    try:
        data         = request.get_json()
        top_disease  = data.get("top_disease", "Unknown")
        total_cases  = data.get("total_cases", 0)
        peak_month   = data.get("peak_month", "Unknown")
        accuracy     = data.get("accuracy", 0)
        distribution = data.get("distribution", [])
        date_range   = data.get("date_range", "")

        dist_text = "\n".join(
            f"  - {d['metric']}: {d['total']:,} cases ({d['pct']}%)"
            for d in distribution
        )

        prompt = f"""You are a senior medical data analyst reviewing hospital disease surveillance data for Kenya.

Summary:
- Date range: {date_range}
- Total cases: {total_cases}
- Highest disease: {top_disease}
- Peak month: {peak_month}
- Model accuracy: {accuracy}%

Disease distribution:
{dist_text}

Provide a concise analysis with:
1. **Key Findings**
2. **Risk Alerts**
3. **Recommended Actions**
4. **Forecast Outlook**

Be specific and direct."""

        client  = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
        message = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}]
        )

        return jsonify({"success": True, "insights": message.content[0].text})

    except anthropic.AuthenticationError:
        return jsonify({"success": False, "error": "Invalid API key."}), 401
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


# -------------------------------------------------
# DASHBOARD
# -------------------------------------------------

@app.route("/", methods=["GET", "POST"])
@login_required
def dashboard():

    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            flash("No file selected.")
            return redirect(url_for("dashboard"))

        try:
            df = pd.read_csv(file)
            df.columns = df.columns.str.strip()

            date_column = None
            for col in df.columns:
                if "date" in col.lower():
                    date_column = col
                    break

            if not date_column:
                flash("No date column detected.")
                return redirect(url_for("dashboard"))

            df[date_column] = pd.to_datetime(df[date_column], errors="coerce")
            HospitalMetric.query.delete()

            for _, row in df.iterrows():
                for col in df.columns:
                    if col == date_column:
                        continue
                    if pd.api.types.is_numeric_dtype(df[col]):
                        db.session.add(HospitalMetric(
                            date=row[date_column],
                            metric_name=col,
                            metric_value=row[col]
                        ))

            db.session.commit()
            flash("Data uploaded and analyzed successfully.")

        except Exception as e:
            flash(f"Upload error: {str(e)}")

        return redirect(url_for("dashboard"))

    # ── ANALYSIS ──
    records = HospitalMetric.query.all()
    if not records:
        return render_template("dashboard.html", message="No data uploaded yet.")

    df = pd.DataFrame([{
        "date": r.date, "metric": r.metric_name, "value": r.metric_value
    } for r in records])

    total_cases   = int(df["value"].sum())
    metric_totals = df.groupby("metric")["value"].sum().sort_values(ascending=False)
    top_disease   = metric_totals.index[0]

    df["month"] = pd.to_datetime(df["date"]).dt.strftime("%B")
    peak_month     = df.groupby("month")["value"].sum().idxmax()
    peak_month_pct = round(df.groupby("month")["value"].sum().pct_change().max() * 100, 1)

    df_grouped = df.groupby(["date", "metric"])["value"].sum().reset_index()
    df_grouped["MonthIndex"] = df_grouped.groupby("metric").cumcount() + 1

    top_df = df_grouped[df_grouped["metric"] == top_disease].copy()
    X, y   = top_df[["MonthIndex"]], top_df["value"]

    model = RandomForestRegressor(n_estimators=100, random_state=42)
    model.fit(X, y)
    predictions    = model.predict(X)
    accuracy       = round(r2_score(y, predictions) * 100, 1)
    predicted_total = int(sum(predictions))

    case_distribution = [
        {"metric": m, "total": int(t), "pct": round(t / metric_totals.sum() * 100, 1)}
        for m, t in metric_totals.items()
    ]

    colors     = ["#4A9EFF", "#FF8C42", "#4CAF87", "#E94E77", "#9B59B6"]
    trend_fig  = go.Figure()
    for i, metric in enumerate(df_grouped["metric"].unique()):
        mdf = df_grouped[df_grouped["metric"] == metric]
        trend_fig.add_trace(go.Scatter(
            x=mdf["date"].astype(str), y=mdf["value"],
            mode="lines", name=metric,
            line=dict(color=colors[i % len(colors)], width=2)
        ))

    trend_fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#CBD5E1"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=40, r=20, t=40, b=40),
        xaxis=dict(gridcolor="rgba(255,255,255,0.05)", showline=False),
        yaxis=dict(gridcolor="rgba(255,255,255,0.05)", showline=False),
        hovermode="x unified"
    )
    trend_chart = json.dumps(trend_fig, default=str)

    min_date   = df["date"].min()
    max_date   = df["date"].max()
    date_range = f"{pd.to_datetime(min_date).strftime('%b %Y')} – {pd.to_datetime(max_date).strftime('%b %Y')}"

    return render_template(
        "dashboard.html",
        total_cases=f"{total_cases:,}",
        top_disease=top_disease,
        peak_month=peak_month,
        peak_month_pct=peak_month_pct,
        accuracy=accuracy,
        predicted_total=f"{predicted_total:,}",
        case_distribution=case_distribution,
        trend_chart=trend_chart,
        date_range=date_range,
        username=current_user.username
    )


# -------------------------------------------------
# INITIALIZE DATABASE
# -------------------------------------------------

with app.app_context():
    db.create_all()

# -------------------------------------------------
# RUN
# -------------------------------------------------

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)