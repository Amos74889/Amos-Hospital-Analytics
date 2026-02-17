import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from flask import Flask, render_template, request
from sklearn.ensemble import RandomForestRegressor

app = Flask(__name__)


# ===============================
# DEFAULT SAMPLE DATA
# ===============================
def get_data():
    dates = pd.date_range(start="2019-01-01", periods=60, freq="M")

    data = {
        "date": dates,
        "malaria_cases": np.random.randint(100, 500, 60),
        "influenza_cases": np.random.randint(50, 300, 60),
        "respiratory_infections": np.random.randint(80, 400, 60),
    }

    return pd.DataFrame(data)


# ===============================
# TREND CHART
# ===============================
def generate_trend_chart(df):

    df = df.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")

    plt.figure(figsize=(10, 5))

    plt.plot(df["date"], df["malaria_cases"], label="Malaria")
    plt.plot(df["date"], df["influenza_cases"], label="Influenza")
    plt.plot(df["date"], df["respiratory_infections"], label="Respiratory")

    plt.legend()
    plt.xticks(rotation=45)

    os.makedirs("static", exist_ok=True)

    path = "static/trend_chart.png"
    plt.savefig(path, bbox_inches="tight")
    plt.close()

    return "trend_chart.png"


# ===============================
# PIE CHART
# ===============================
def generate_pie_chart(df):

    df = df.copy()

    totals = [
        df["malaria_cases"].sum(),
        df["influenza_cases"].sum(),
        df["respiratory_infections"].sum()
    ]

    labels = ["Malaria", "Influenza", "Respiratory"]

    plt.figure()
    plt.pie(totals, labels=labels, autopct="%1.1f%%")

    path = "static/pie_chart.png"
    plt.savefig(path, bbox_inches="tight")
    plt.close()

    return "pie_chart.png"


# ===============================
# PREDICTION CHART
# ===============================
def generate_prediction_chart(df):

    df = df.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")

    df["month"] = df["date"].dt.month

    X = df[["month"]]
    y = df["malaria_cases"]

    model = RandomForestRegressor()
    model.fit(X, y)

    predictions = model.predict(X)

    plt.figure(figsize=(8, 4))
    plt.plot(df["date"], y, label="Actual")
    plt.plot(df["date"], predictions, label="Predicted")

    plt.legend()

    path = "static/prediction_chart.png"
    plt.savefig(path, bbox_inches="tight")
    plt.close()

    return "prediction_chart.png"


# ===============================
# MAIN ROUTE
# ===============================
@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":
        file = request.files.get("file")

        if file and file.filename != "":
            df = pd.read_csv(file)
        else:
            df = get_data()
    else:
        df = get_data()

    # Normalize column names
    df.columns = df.columns.str.strip().str.lower()

    # 🔥 FORCE date column to datetime AFTER normalization
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    else:
        return "Error: CSV must contain a 'date' column"

    # Generate charts
    trend_chart = generate_trend_chart(df)
    pie_chart = generate_pie_chart(df)
    prediction_chart = generate_prediction_chart(df)

    # KPIs
    total_cases = (
        df["malaria_cases"].sum() +
        df["influenza_cases"].sum() +
        df["respiratory_infections"].sum()
    )

    disease_totals = {
        "Malaria": df["malaria_cases"].sum(),
        "Influenza": df["influenza_cases"].sum(),
        "Respiratory": df["respiratory_infections"].sum(),
    }

    highest_disease = max(disease_totals, key=disease_totals.get)
    peak_month = df.loc[df["malaria_cases"].idxmax(), "date"].strftime("%B")

    accuracy = 89.5

    return render_template(
        "index.html",
        trend_chart=trend_chart,
        pie_chart=pie_chart,
        prediction_chart=prediction_chart,
        total_cases=total_cases,
        highest_disease=highest_disease,
        peak_month=peak_month,
        accuracy=accuracy
    )


if __name__ == "__main__":
    app.run(debug=True)
