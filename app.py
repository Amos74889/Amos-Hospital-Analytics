import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager,
    UserMixin,
    login_user,
    login_required,
    logout_user,
    current_user
)
from werkzeug.security import generate_password_hash, check_password_hash
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import r2_score
import plotly.express as px

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


# ✅ DYNAMIC METRIC MODEL
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
# DASHBOARD (DYNAMIC ENGINE)
# -------------------------------------------------

@app.route("/", methods=["GET", "POST"])
@login_required
def dashboard():

    # ---------------- FILE UPLOAD ----------------
    if request.method == "POST":
        file = request.files.get("file")

        if not file:
            flash("No file selected.")
            return redirect(url_for("dashboard"))

        try:
            df = pd.read_csv(file)
            df.columns = df.columns.str.strip()

            # Detect date column automatically
            date_column = None
            for col in df.columns:
                if "date" in col.lower():
                    date_column = col
                    break

            if not date_column:
                flash("No date column detected.")
                return redirect(url_for("dashboard"))

            df[date_column] = pd.to_datetime(df[date_column], errors="coerce")

            # Clear old data
            HospitalMetric.query.delete()

            # Store dynamically
            for _, row in df.iterrows():
                for col in df.columns:
                    if col == date_column:
                        continue

                    if pd.api.types.is_numeric_dtype(df[col]):
                        metric = HospitalMetric(
                            date=row[date_column],
                            metric_name=col,
                            metric_value=row[col]
                        )
                        db.session.add(metric)

            db.session.commit()
            flash("Hospital data uploaded and analyzed successfully.")

        except Exception as e:
            flash(f"Upload error: {str(e)}")

        return redirect(url_for("dashboard"))

    # ---------------- ANALYSIS ----------------

    records = HospitalMetric.query.all()

    if not records:
        return render_template("dashboard.html", message="No data uploaded yet.")

    df = pd.DataFrame([{
        "date": r.date,
        "metric": r.metric_name,
        "value": r.metric_value
    } for r in records])

    total_cases = int(df["value"].sum())

    top_metric = (
        df.groupby("metric")["value"]
        .sum()
        .sort_values(ascending=False)
        .index[0]
    )

    # Prepare grouped data
    df_grouped = df.groupby(["date", "metric"])["value"].sum().reset_index()
    df_grouped["MonthIndex"] = df_grouped.groupby("metric").cumcount() + 1

    # ML forecast only top metric
    top_df = df_grouped[df_grouped["metric"] == top_metric]

    X = top_df[["MonthIndex"]]
    y = top_df["value"]

    model = RandomForestRegressor(n_estimators=100, random_state=42)
    model.fit(X, y)

    predictions = model.predict(X)
    accuracy = round(r2_score(y, predictions) * 100, 2)

    future_df = pd.DataFrame({
        "MonthIndex": range(len(top_df) + 1, len(top_df) + 7)
    })

    forecast = model.predict(future_df)

    # Trend chart
    trend_fig = px.line(
        df_grouped,
        x="date",
        y="value",
        color="metric",
        title="Hospital Metrics Trend"
    )
    trend_chart = trend_fig.to_html(full_html=False)

    # Forecast chart
    forecast_fig = px.line(
        x=list(range(1, 7)),
        y=forecast,
        labels={"x": "Next 6 Months", "y": f"Forecasted {top_metric}"},
        title=f"6-Month Forecast for {top_metric}"
    )
    forecast_chart = forecast_fig.to_html(full_html=False)

    return render_template(
        "dashboard.html",
        total_cases=total_cases,
        top_disease=top_metric,
        accuracy=accuracy,
        trend_chart=trend_chart,
        forecast_chart=forecast_chart
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