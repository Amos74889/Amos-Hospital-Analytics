from flask import Flask, render_template, request
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio

from models import load_data, train_random_forest, train_arima
from config import DevelopmentConfig

app = Flask(__name__)
app.config.from_object(DevelopmentConfig)

# Load dataset once
data = load_data()


def generate_trend_chart(df):

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df["Date"],
        y=df["Malaria_Cases"],
        mode="lines",
        name="Malaria",
        line=dict(width=3)
    ))

    fig.add_trace(go.Scatter(
        x=df["Date"],
        y=df["Influenza_Cases"],
        mode="lines",
        name="Influenza",
        line=dict(width=3)
    ))

    fig.add_trace(go.Scatter(
        x=df["Date"],
        y=df["Respiratory_Infections"],
        mode="lines",
        name="Respiratory Infections",
        line=dict(width=3)
    ))

    fig.update_layout(
        template="plotly_white",
        title="Disease Trends in Kenya",
        xaxis_title="Date",
        yaxis_title="Cases",
        hovermode="x unified",
        legend=dict(orientation="h"),
        margin=dict(l=40, r=40, t=60, b=40)
    )

    return pio.to_html(fig, full_html=False)


def generate_pie_chart(df):

    totals = [
        df["Malaria_Cases"].sum(),
        df["Influenza_Cases"].sum(),
        df["Respiratory_Infections"].sum()
    ]

    fig = go.Figure(data=[go.Pie(
        labels=["Malaria", "Influenza", "Respiratory"],
        values=totals,
        hole=0.5
    )])

    fig.update_layout(
        template="plotly_white",
        title="Case Distribution"
    )

    return pio.to_html(fig, full_html=False)


@app.route("/", methods=["GET", "POST"])
def index():

    results = {}

    if request.method == "POST":
        file = request.files.get("file")

        if file:
            df = pd.read_csv(file)
        else:
            df = data
    else:
        df = data

    # Generate interactive charts
    trend_chart = generate_trend_chart(df)
    pie_chart = generate_pie_chart(df)

    # Train models only when POST
    if request.method == "POST":
        rf_result = train_random_forest(df)
        arima_model = train_arima(df)

        results = {
            "rf_mae": rf_result.get("MAE", "N/A"),
            "rf_rmse": rf_result.get("RMSE", "N/A"),
            "arima_summary": arima_model.summary().as_text()
        }

    return render_template(
        "index.html",
        results=results,
        trend_chart=trend_chart,
        pie_chart=pie_chart
    )


if __name__ == "__main__":
    app.run(debug=app.config["DEBUG"])
