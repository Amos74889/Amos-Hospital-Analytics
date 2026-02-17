import os
import numpy as np
import pandas as pd
from flask import Flask, render_template
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import r2_score
import plotly.express as px
import plotly.graph_objects as go

app = Flask(__name__)

# -----------------------------
# Generate Sample Hospital Data
# -----------------------------
def generate_data():
    np.random.seed(42)

    # IMPORTANT: Using "ME" (Month End) for new pandas versions
    dates = pd.date_range(start="2022-01-01", periods=24, freq="ME")

    df = pd.DataFrame({
        "Date": dates,
        "Malaria": np.random.randint(500, 1200, 24),
        "Influenza": np.random.randint(200, 800, 24),
        "Respiratory": np.random.randint(400, 1000, 24),
    })

    return df


# -----------------------------
# Main Dashboard Route
# -----------------------------
@app.route("/")
def dashboard():

    df = generate_data()

    # -----------------------------
    # KPI Calculations
    # -----------------------------
    total_cases = int(df[["Malaria", "Influenza", "Respiratory"]].sum().sum())
    disease_totals = df[["Malaria", "Influenza", "Respiratory"]].sum()
    highest_disease = disease_totals.idxmax()
    peak_month = df.loc[df["Malaria"].idxmax(), "Date"].strftime("%B")

    # -----------------------------
    # Machine Learning Model
    # -----------------------------
    df["Month"] = df["Date"].dt.month

    X = df[["Month"]]
    y = df["Malaria"]

    model = RandomForestRegressor(n_estimators=50, random_state=42)
    model.fit(X, y)

    predictions = model.predict(X)
    accuracy = round(r2_score(y, predictions) * 100, 2)

    # -----------------------------
    # Charts
    # -----------------------------

    # Trend Chart
    trend_fig = px.line(
        df,
        x="Date",
        y=["Malaria", "Influenza", "Respiratory"],
        markers=True,
        title="Disease Trends Over Time"
    )
    trend_fig.update_layout(template="plotly_white")
    trend_chart = trend_fig.to_html(full_html=False, include_plotlyjs=False)

    # Pie Chart
    pie_fig = px.pie(
        names=disease_totals.index,
        values=disease_totals.values,
        hole=0.5,
        title="Disease Distribution"
    )
    pie_fig.update_layout(template="plotly_white")
    pie_chart = pie_fig.to_html(full_html=False, include_plotlyjs=False)

    # Prediction Chart
    pred_fig = go.Figure()
    pred_fig.add_trace(go.Scatter(
        x=df["Date"],
        y=y,
        mode="lines+markers",
        name="Actual"
    ))
    pred_fig.add_trace(go.Scatter(
        x=df["Date"],
        y=predictions,
        mode="lines+markers",
        name="Predicted"
    ))
    pred_fig.update_layout(
        title="Malaria Prediction vs Actual",
        template="plotly_white"
    )
    prediction_chart = pred_fig.to_html(full_html=False, include_plotlyjs=False)

    # Data Table
    table_html = df.to_html(
        classes="table table-hover table-striped",
        index=False
    )

    # -----------------------------
    # Render Template
    # -----------------------------
    return render_template(
        "index.html",
        total_cases=f"{total_cases:,}",
        highest_disease=highest_disease,
        peak_month=peak_month,
        accuracy=accuracy,
        trend_chart=trend_chart,
        pie_chart=pie_chart,
        prediction_chart=prediction_chart,
        data_table=table_html
    )


# -----------------------------
# Run Application
# -----------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
