import os
import io
import datetime
import pandas as pd
import numpy as np
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, UserMixin, login_user,
    login_required, logout_user, current_user
)
from werkzeug.security import generate_password_hash, check_password_hash
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import r2_score, mean_absolute_error, mean_squared_error
from sklearn.model_selection import train_test_split
from statsmodels.tsa.arima.model import ARIMA
import plotly.graph_objects as go
import json
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from dotenv import load_dotenv

load_dotenv()

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
# HELPER: RUN RANDOM FOREST
# -------------------------------------------------

def run_random_forest(df_grouped, top_disease):
    top_df = df_grouped[df_grouped["metric"] == top_disease].copy()
    top_df["MonthIndex"] = range(1, len(top_df) + 1)

    X = top_df[["MonthIndex"]]
    y = top_df["value"]

    if len(top_df) < 5:
        predictions = y.values
        return {
            "rf_mae": 0, "rf_rmse": 0, "rf_r2": 100,
            "predicted_total": int(y.sum()),
            "predictions": predictions.tolist()
        }

    # Train on all data for better accuracy with small datasets
    model = RandomForestRegressor(n_estimators=100, random_state=42)
    model.fit(X, y)

    all_preds = model.predict(X)

    # Calculate metrics on full data
    mae  = round(mean_absolute_error(y, all_preds), 2)
    rmse = round(np.sqrt(mean_squared_error(y, all_preds)), 2)
    r2   = round(r2_score(y, all_preds) * 100, 1)

    return {
        "rf_mae": mae,
        "rf_rmse": rmse,
        "rf_r2": max(r2, 0),
        "predicted_total": int(sum(all_preds)),
        "predictions": all_preds.tolist()
    }


# -------------------------------------------------
# HELPER: RUN ARIMA
# -------------------------------------------------

def run_arima(df_grouped, top_disease):
    top_df = df_grouped[df_grouped["metric"] == top_disease].copy()
    series = top_df["value"].values

    if len(series) < 6:
        return {
            "arima_mae": 0, "arima_rmse": 0,
            "arima_forecast": [], "arima_status": "Not enough data"
        }

    try:
        train_size = int(len(series) * 0.8)
        train, test = series[:train_size], series[train_size:]

        model     = ARIMA(train, order=(2, 1, 2))
        model_fit = model.fit()

        forecast  = model_fit.forecast(steps=len(test))

        mae  = round(mean_absolute_error(test, forecast), 2)
        rmse = round(np.sqrt(mean_squared_error(test, forecast)), 2)

        # Forecast next 3 months
        future_model = ARIMA(series, order=(2, 1, 2))
        future_fit   = future_model.fit()
        future_fc    = future_fit.forecast(steps=3)

        return {
            "arima_mae": mae,
            "arima_rmse": rmse,
            "arima_forecast": [round(x, 1) for x in future_fc.tolist()],
            "arima_status": "Success"
        }

    except Exception as e:
        return {
            "arima_mae": 0, "arima_rmse": 0,
            "arima_forecast": [], "arima_status": f"Error: {str(e)}"
        }


# -------------------------------------------------
# HELPER: SURGE ALERTS
# -------------------------------------------------

def detect_surges(df, case_distribution):
    alerts = []

    for item in case_distribution:
        metric = item["metric"]
        metric_df = df[df["metric"] == metric].copy()
        metric_df = metric_df.sort_values("date")

        if len(metric_df) < 2:
            continue

        values = metric_df["value"].values
        mean_val = np.mean(values)
        std_val  = np.std(values)
        last_val = values[-1]

        # Surge if last value is more than 1.5 std above mean
        if last_val > mean_val + 1.5 * std_val:
            pct_above = round(((last_val - mean_val) / mean_val) * 100, 1)
            alerts.append({
                "type": "danger",
                "metric": metric,
                "message": f"⚠️ SURGE ALERT: {metric} is {pct_above}% above average!",
                "value": int(last_val),
                "mean": int(mean_val)
            })
        elif last_val > mean_val + 0.8 * std_val:
            alerts.append({
                "type": "warning",
                "metric": metric,
                "message": f"⚡ WARNING: {metric} is trending above normal levels.",
                "value": int(last_val),
                "mean": int(mean_val)
            })

    return alerts


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
    peak_month   = df.groupby("month")["value"].sum().idxmax()

    df_grouped = df.groupby(["date", "metric"])["value"].sum().reset_index()

    rf_results    = run_random_forest(df_grouped, top_disease)
    arima_results = run_arima(df_grouped, top_disease)

    case_distribution = []
    for metric, total in metric_totals.items():
        pct = round(total / metric_totals.sum() * 100, 1)
        case_distribution.append({"metric": metric, "total": int(total), "pct": pct})

    min_date   = df["date"].min()
    max_date   = df["date"].max()
    date_range = f"{pd.to_datetime(min_date).strftime('%b %Y')} – {pd.to_datetime(max_date).strftime('%b %Y')}"

    alerts = detect_surges(df, case_distribution)

    return {
        "total_cases":       total_cases,
        "top_disease":       top_disease,
        "peak_month":        peak_month,
        "case_distribution": case_distribution,
        "date_range":        date_range,
        "alerts":            alerts,
        # Random Forest metrics
        "rf_mae":            rf_results["rf_mae"],
        "rf_rmse":           rf_results["rf_rmse"],
        "rf_r2":             rf_results["rf_r2"],
        "predicted_total":   rf_results["predicted_total"],
        # ARIMA metrics
        "arima_mae":         arima_results["arima_mae"],
        "arima_rmse":        arima_results["arima_rmse"],
        "arima_forecast":    arima_results["arima_forecast"],
        "arima_status":      arima_results["arima_status"],
    }


# -------------------------------------------------
# HELPER: CALL CLAUDE FOR ANALYSIS
# -------------------------------------------------


# -------------------------------------------------
# HELPER: BUILD SMART CONTEXT-AWARE SYSTEM PROMPT
# -------------------------------------------------

def build_smart_prompt(data, is_full_report=False):
    """Builds a context-aware prompt based on what is happening in the data."""

    # Extract values
    alerts       = data.get("alerts", [])
    arima_fc     = data.get("arima_forecast", [])
    top_disease  = data.get("top_disease", "Unknown")
    rf_r2        = data.get("rf_r2", 0)
    arima_mae    = data.get("arima_mae", 0)
    rf_mae       = data.get("rf_mae", 0)
    total_cases  = data.get("total_cases", 0)
    peak_month   = data.get("peak_month", "Unknown")
    date_range   = data.get("date_range", "")
    distribution = data.get("distribution", [])
    rf_rmse      = data.get("rf_rmse", 0)
    arima_rmse   = data.get("arima_rmse", 0)

    dist_text    = "\n".join(f"  - {d['metric']}: {d['total']:,} cases ({d['pct']}%)" for d in distribution)
    alerts_text  = "\n".join(f"  - {a['message']}" for a in alerts) if alerts else "  - No active surge alerts"

    # ── SITUATION DETECTION ──
    has_surge        = any(a["type"] == "danger" for a in alerts)
    has_warning      = any(a["type"] == "warning" for a in alerts)
    forecast_rising  = len(arima_fc) >= 2 and arima_fc[-1] > arima_fc[0]
    forecast_falling = len(arima_fc) >= 2 and arima_fc[-1] < arima_fc[0]
    low_accuracy     = rf_r2 < 60
    high_accuracy    = rf_r2 >= 85

    # ── BUILD SITUATIONAL RULES ──
    situation_rules = []

    if has_surge:
        situation_rules.append("""
⚠️ SURGE ALERT ACTIVE — You MUST include these specific recommendations:
- STAFF: Immediately increase on-call staff by at least 20-30% in affected wards
- BEDS: Activate overflow bed protocols and prepare emergency ward expansion
- SUPPLIES: Issue urgent procurement order for medicines related to """ + top_disease + """
- MANAGEMENT: Implement triage fast-track system to reduce waiting times
- ESCALATION: Notify hospital administration and county health department immediately""")

    if has_warning:
        situation_rules.append("""
⚡ WARNING LEVEL ALERT — You MUST include these precautionary recommendations:
- STAFF: Place additional staff on standby rotation for the next 2 weeks
- MONITORING: Increase data collection frequency from monthly to weekly
- SUPPLIES: Pre-order a 30-day buffer stock of critical medicines
- MANAGEMENT: Brief ward managers on potential surge preparedness""")

    if forecast_rising:
        fc_vals = " → ".join(str(int(f)) for f in arima_fc)
        situation_rules.append(f"""
📈 ARIMA FORECAST IS RISING ({fc_vals}) — You MUST warn administrators:
- BEDS: Plan for {int(arima_fc[-1] * 1.1):,} cases next month — begin bed allocation planning now
- STAFF: Schedule additional staff rotations starting next month
- MEDICINE: Order medicines 3 weeks in advance based on forecast of {int(arima_fc[-1]):,} cases
- MANAGEMENT: Set up a weekly review meeting to track if forecast is accurate""")

    if forecast_falling:
        fc_vals = " → ".join(str(int(f)) for f in arima_fc)
        situation_rules.append(f"""
📉 ARIMA FORECAST IS DECLINING ({fc_vals}) — You MUST note this positive trend:
- STAFF: Maintain current staffing but avoid new hires until trend is confirmed
- SUPPLIES: Reduce next procurement order by 15-20% to avoid waste
- MANAGEMENT: Use this period to conduct staff training and equipment maintenance
- DATA: Investigate what interventions caused the decline and document them""")

    if low_accuracy:
        situation_rules.append(f"""
🔴 MODEL ACCURACY IS LOW (R²: {rf_r2}%) — You MUST recommend data improvements:
- DATA COLLECTION: Current dataset is insufficient — collect at minimum 24 months of data
- VARIABLES: Add weather data, population density, and seasonal disease patterns
- FREQUENCY: Switch from monthly to weekly data collection for better model training
- VALIDATION: Consider hiring a data officer to ensure data quality and completeness
- MODELS: With more data, try XGBoost or LSTM neural networks for better accuracy""")

    if high_accuracy:
        situation_rules.append(f"""
✅ MODEL ACCURACY IS HIGH (R²: {rf_r2}%) — Acknowledge this and build on it:
- The predictive system is performing well and can be trusted for planning decisions
- Recommend expanding the model to predict individual ward-level admissions
- Suggest sharing this model with other hospitals in the region""")

    # ── DISEASE-SPECIFIC RULES ──
    disease_rules = {
        "malaria": """
🦟 MALARIA IS THE TOP DISEASE — Specific actions required:
- MEDICINE: Ensure adequate stock of Artemether-Lumefantrine (AL) and RDT test kits
- STAFF: Train nurses on updated malaria case management protocols
- PREVENTION: Coordinate with county government for mosquito net distribution campaigns
- SEASONAL: Malaria peaks March-May and October-November in Kenya — plan accordingly""",

        "dengue": """
🦟 DENGUE IS THE TOP DISEASE — Specific actions required:
- MEDICINE: Stock up on paracetamol, IV fluids, and platelet monitoring equipment
- STAFF: Ensure staff are trained to recognize dengue hemorrhagic fever warning signs
- PREVENTION: Report to Kenya CDC for vector control fumigation in surrounding areas
- MONITORING: Monitor platelet counts daily for all dengue patients""",

        "cholera": """
💧 CHOLERA IS THE TOP DISEASE — Specific actions required:
- MEDICINE: Urgently stock ORS (Oral Rehydration Salts) and IV Ringers Lactate
- ISOLATION: Set up dedicated cholera treatment unit away from main wards
- WATER: Report to county water authority immediately — likely contaminated water source
- STAFF: All staff must use PPE — cholera spreads rapidly in hospital settings
- REPORTING: This is a notifiable disease — report to Ministry of Health within 24 hours""",

        "icu": """
🏥 ICU ADMISSIONS ARE HIGH — Specific actions required:
- BEDS: Review ICU bed capacity immediately — consider converting HDU beds
- STAFF: Ensure ICU nurse-to-patient ratio does not exceed 1:2
- EQUIPMENT: Check availability of ventilators, monitors, and infusion pumps
- REFERRALS: Establish referral protocol with Kenyatta National Hospital for overflow""",
    }

    # Match disease
    disease_advice = ""
    for key, advice in disease_rules.items():
        if key.lower() in top_disease.lower():
            disease_advice = advice
            break

    if not disease_advice:
        disease_advice = f"""
🔬 {top_disease} IS THE TOP DISEASE — Specific actions required:
- MEDICINE: Audit current stock levels for {top_disease} treatment protocols
- STAFF: Brief clinical staff on current case management guidelines
- MONITORING: Increase monitoring frequency for {top_disease} patients
- REPORTING: Ensure all {top_disease} cases are properly coded and reported"""

    # ── ASSEMBLE FULL PROMPT ──
    situation_block = "\n".join(situation_rules) if situation_rules else "  - All metrics within normal range"

    sections = """1. **Executive Summary**
2. **Key Findings**
3. **Disease Burden Analysis** (focus on """ + top_disease + """)
4. **Risk Alerts & Surge Detection**
5. **Staff & Resource Allocation** (give specific numbers)
6. **Medicine & Supply Orders** (name specific medicines)
7. **Patient Management Advice**
8. **Model Performance** (Random Forest vs ARIMA comparison)
9. **Data Collection Improvements**
10. **3-Month Forecast Outlook**""" if is_full_report else """1. **Key Findings**
2. **Risk Alerts & Surge Detection**
3. **Staff & Resource Allocation** (give specific numbers)
4. **Medicine & Supply Orders** (name specific medicines)
5. **Patient Management Advice**
6. **Model Comparison** (Random Forest vs ARIMA)
7. **Data Collection Improvements**
8. **Forecast Outlook**"""

    prompt = f"""You are a senior medical data analyst and hospital resource planning expert for Kenya's healthcare system.

=== HOSPITAL DATA ===
- Date range: {date_range}
- Total cases: {total_cases:,}
- Highest disease burden: {top_disease}
- Peak admission month: {peak_month}

=== MODEL PERFORMANCE ===
- Random Forest — MAE: {rf_mae}, RMSE: {rf_rmse}, R²: {rf_r2}%
- ARIMA — MAE: {arima_mae}, RMSE: {arima_rmse}
- ARIMA 3-Month Forecast: {arima_fc}

=== ACTIVE ALERTS ===
{alerts_text}

=== DISEASE DISTRIBUTION ===
{dist_text}

=== SITUATIONAL RULES — YOU MUST FOLLOW THESE ===
{situation_block}

=== DISEASE-SPECIFIC GUIDANCE ===
{disease_advice}

=== YOUR TASK ===
Write a professional, actionable hospital analytics report with these sections:
{sections}

RULES:
- Be SPECIFIC — use actual numbers from the data
- Name SPECIFIC medicines, not just "medications"
- Give SPECIFIC staff numbers (e.g. "add 3 nurses per shift")
- Reference SPECIFIC Kenyan context (Kenya CDC, Ministry of Health, county health departments)
- If a surge is detected, treat it as URGENT and say so clearly
- If forecast is rising, give a CLEAR WARNING with specific preparation steps
- Keep each section concise but actionable"""

    return prompt


def get_ai_analysis(summary):
    dist_text = "\n".join(
        f"  - {d['metric']}: {d['total']:,} cases ({d['pct']}%)"
        for d in summary["case_distribution"]
    )

    alerts_text = "\n".join(
        f"  - {a['message']}" for a in summary["alerts"]
    ) if summary["alerts"] else "  - No active surge alerts"

    # Build context-aware data dict for smart prompt
    summary_data = {
        "alerts":       summary["alerts"],
        "arima_forecast": summary["arima_forecast"],
        "top_disease":  summary["top_disease"],
        "rf_r2":        summary["rf_r2"],
        "rf_mae":       summary["rf_mae"],
        "rf_rmse":      summary["rf_rmse"],
        "arima_mae":    summary["arima_mae"],
        "arima_rmse":   summary["arima_rmse"],
        "total_cases":  summary["total_cases"],
        "peak_month":   summary["peak_month"],
        "date_range":   summary["date_range"],
        "distribution": summary["case_distribution"],
    }
    prompt = build_smart_prompt(summary_data, is_full_report=True)

    genai.configure(api_key=os.environ.get("GEMINI_API_KEY"))
    model  = genai.GenerativeModel("gemini-2.0-flash")
    result = model.generate_content(prompt)
    return result.text


# -------------------------------------------------
# HELPER: BUILD WORD DOCUMENT
# -------------------------------------------------

def build_word_report(summary, ai_text):
    doc = Document()

    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.25)
        section.right_margin  = Inches(1.25)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run("AMOS HOSPITAL ANALYTICS")
    run.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(0x0D, 0x47, 0xA1)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = sub_p.add_run("Disease Surveillance & Predictive Analytics Report")
    run2.font.size = Pt(14)
    run2.font.color.rgb = RGBColor(0x37, 0x47, 0x5A)

    doc.add_paragraph()

    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_p.add_run(f"Report Period: {summary['date_range']}").bold = True

    gen_p = doc.add_paragraph()
    gen_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    gen_p.add_run(f"Generated: {datetime.datetime.now().strftime('%B %d, %Y at %H:%M')}")

    doc.add_paragraph()

    # KPI Table
    h = doc.add_heading("Summary Statistics", level=1)
    h.runs[0].font.color.rgb = RGBColor(0x0D, 0x47, 0xA1)

    table  = doc.add_table(rows=2, cols=4)
    table.style = "Table Grid"
    headers = ["Total Cases", "Highest Disease", "Peak Month", "Predicted Cases"]
    values  = [
        f"{summary['total_cases']:,}",
        summary["top_disease"],
        summary["peak_month"],
        f"{summary['predicted_total']:,}"
    ]

    for i, (hdr, val) in enumerate(zip(headers, values)):
        cell = table.rows[0].cells[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        hr = cell.paragraphs[0].add_run(hdr)
        hr.bold = True; hr.font.size = Pt(9)
        hr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        tc = cell._tc; tcp = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "1565C0"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:val"), "clear")
        tcp.append(shd)

        vc = table.rows[1].cells[i]
        vc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        vr = vc.paragraphs[0].add_run(val)
        vr.bold = True; vr.font.size = Pt(13)

    doc.add_paragraph()

    # Model Metrics Table
    h2 = doc.add_heading("Model Performance Metrics", level=1)
    h2.runs[0].font.color.rgb = RGBColor(0x0D, 0x47, 0xA1)

    mtable = doc.add_table(rows=1, cols=4)
    mtable.style = "Table Grid"
    mhdrs  = ["Model", "MAE", "RMSE", "R² / Status"]
    mvals  = [
        ("Random Forest", str(summary["rf_mae"]), str(summary["rf_rmse"]), f"{summary['rf_r2']}%"),
        ("ARIMA", str(summary["arima_mae"]), str(summary["arima_rmse"]), summary["arima_status"]),
    ]

    for i, mh in enumerate(mhdrs):
        cell = mtable.rows[0].cells[i]
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = cell.paragraphs[0].add_run(mh)
        r.bold = True; r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        tc = cell._tc; tcp = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "1565C0"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:val"), "clear")
        tcp.append(shd)

    for row_data in mvals:
        row = mtable.add_row().cells
        for ci, val in enumerate(row_data):
            row[ci].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row[ci].paragraphs[0].add_run(val).font.size = Pt(10)

    doc.add_paragraph()

    # Alerts
    if summary["alerts"]:
        h3 = doc.add_heading("Surge Alerts", level=1)
        h3.runs[0].font.color.rgb = RGBColor(0xC6, 0x28, 0x28)
        for alert in summary["alerts"]:
            p = doc.add_paragraph()
            r = p.add_run(alert["message"])
            r.bold = True
            r.font.color.rgb = RGBColor(0xC6, 0x28, 0x28) if alert["type"] == "danger" else RGBColor(0xE6, 0x5C, 0x00)

    doc.add_paragraph()
    doc.add_page_break()

    # AI Analysis
    h4 = doc.add_heading("AI-Powered Analysis & Recommendations", level=1)
    h4.runs[0].font.color.rgb = RGBColor(0x0D, 0x47, 0xA1)

    for line in ai_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        is_heading = (
            (line[0].isdigit() and ". " in line[:4]) or
            line.startswith("##") or
            (line.startswith("**") and line.endswith("**"))
        )
        if is_heading:
            clean = line.lstrip("#").lstrip("0123456789.").strip().strip("*").strip()
            h = doc.add_heading(clean, level=2)
            h.runs[0].font.color.rgb = RGBColor(0x15, 0x65, 0xC0)
        elif line.startswith("- ") or line.startswith("• "):
            p = doc.add_paragraph(style="List Bullet")
            p.add_run(line.lstrip("- ").lstrip("• ")).font.size = Pt(11)
        else:
            p = doc.add_paragraph()
            p.add_run(line).font.size = Pt(11)

    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = footer_p.add_run(
        f"Confidential — Amos Hospital Analytics · Generated {datetime.datetime.now().strftime('%Y-%m-%d')}"
    )
    fr.font.size = Pt(9)
    fr.font.color.rgb = RGBColor(0x9E, 0x9E, 0x9E)
    fr.italic = True

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

        ai_text = get_ai_analysis(summary)
        doc_buf = build_word_report(summary, ai_text)
        filename = f"Amos_Hospital_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.docx"

        return send_file(
            doc_buf,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as auth_err:
        if "API_KEY" in str(auth_err).upper() or "auth" in str(auth_err).lower():
            flash("Invalid or missing Gemini API key.")
            return redirect(url_for("dashboard"))
    except Exception as e:
        flash(f"Report error: {str(e)}")
        return redirect(url_for("dashboard"))


# -------------------------------------------------
# AI INSIGHTS ROUTE
# -------------------------------------------------

@app.route("/ai-insights", methods=["POST"])
@login_required
def ai_insights():
    try:
        data         = request.get_json()
        top_disease  = data.get("top_disease", "Unknown")
        total_cases  = data.get("total_cases", 0)
        peak_month   = data.get("peak_month", "Unknown")
        rf_mae       = data.get("rf_mae", 0)
        rf_rmse      = data.get("rf_rmse", 0)
        rf_r2        = data.get("rf_r2", 0)
        arima_mae    = data.get("arima_mae", 0)
        arima_rmse   = data.get("arima_rmse", 0)
        arima_fc     = data.get("arima_forecast", [])
        distribution = data.get("distribution", [])
        alerts       = data.get("alerts", [])
        date_range   = data.get("date_range", "")

        dist_text   = "\n".join(f"  - {d['metric']}: {d['total']:,} cases ({d['pct']}%)" for d in distribution)
        alerts_text = "\n".join(f"  - {a['message']}" for a in alerts) if alerts else "  - No active surge alerts"

        # Build context-aware data dict
        smart_data = {
            "alerts":         alerts,
            "arima_forecast": arima_fc,
            "top_disease":    top_disease,
            "rf_r2":          rf_r2,
            "rf_mae":         rf_mae,
            "rf_rmse":        rf_rmse,
            "arima_mae":      arima_mae,
            "arima_rmse":     arima_rmse,
            "total_cases":    total_cases,
            "peak_month":     peak_month,
            "date_range":     date_range,
            "distribution":   distribution,
        }
        prompt = build_smart_prompt(smart_data, is_full_report=False)

        genai.configure(api_key=os.environ.get("GEMINI_API_KEY"))
        model  = genai.GenerativeModel("gemini-2.0-flash")
        result = model.generate_content(prompt)

        return jsonify({"success": True, "insights": result.text})

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

    df["month"]    = pd.to_datetime(df["date"]).dt.strftime("%B")
    peak_month     = df.groupby("month")["value"].sum().idxmax()
    peak_month_pct = round(df.groupby("month")["value"].sum().pct_change().max() * 100, 1)

    df_grouped = df.groupby(["date", "metric"])["value"].sum().reset_index()

    # Run both models
    rf_results    = run_random_forest(df_grouped, top_disease)
    arima_results = run_arima(df_grouped, top_disease)

    case_distribution = [
        {"metric": m, "total": int(t), "pct": round(t / metric_totals.sum() * 100, 1)}
        for m, t in metric_totals.items()
    ]

    alerts = detect_surges(df, case_distribution)

    # Trend chart
    colors    = ["#4A9EFF", "#FF8C42", "#4CAF87", "#E94E77", "#9B59B6"]
    trend_fig = go.Figure()
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
        # Random Forest
        rf_mae=rf_results["rf_mae"],
        rf_rmse=rf_results["rf_rmse"],
        rf_r2=rf_results["rf_r2"],
        predicted_total=f"{rf_results['predicted_total']:,}",
        # ARIMA
        arima_mae=arima_results["arima_mae"],
        arima_rmse=arima_results["arima_rmse"],
        arima_forecast=arima_results["arima_forecast"],
        arima_status=arima_results["arima_status"],
        # Other
        case_distribution=case_distribution,
        alerts=alerts,
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

with app.app_context():
    db.create_all()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)