import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_absolute_error, mean_squared_error
from statsmodels.tsa.arima.model import ARIMA
import plotly.graph_objects as go
import plotly.io as pio


# -----------------------------
# 1️⃣ Load Dataset
# -----------------------------
def load_data():
    data = pd.read_csv("data/who_kenya_disease_data.csv")
    data["Date"] = pd.to_datetime(data["Date"])
    data.set_index("Date", inplace=True)
    return data


# -----------------------------
# 2️⃣ Feature Engineering
# -----------------------------
def prepare_features(data):
    data["Hospital_Admissions"] = (
        0.3 * data["Influenza_Cases"] +
        0.4 * data["Malaria_Cases"] +
        0.3 * data["Respiratory_Infections"]
    )

    data["Month"] = data.index.month
    data["Year"] = data.index.year

    return data


# -----------------------------
# 3️⃣ Train Random Forest
# -----------------------------
def train_random_forest(data):
    features = ["Influenza_Cases", "Malaria_Cases",
                "Respiratory_Infections", "Month", "Year"]

    X = data[features]
    y = data["Hospital_Admissions"]

    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, shuffle=False
    )

    model = RandomForestRegressor(n_estimators=200, random_state=42)
    model.fit(X_train, y_train)

    predictions = model.predict(X_test)

    mae = mean_absolute_error(y_test, predictions)
    rmse = np.sqrt(mean_squared_error(y_test, predictions))

    return {
        "MAE": round(mae, 2),
        "RMSE": round(rmse, 2)
    }


# -----------------------------
# 4️⃣ Train ARIMA
# -----------------------------
def train_arima(data):
    series = data["Hospital_Admissions"]
    model = ARIMA(series, order=(2, 1, 2))
    model_fit = model.fit()

    return model_fit.summary().as_text()


# -----------------------------
# 5️⃣ Interactive Line Chart
# -----------------------------
def generate_trend_chart(data):

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=data.index,
        y=data["Malaria_Cases"],
        mode="lines",
        name="Malaria",
        line=dict(width=3)
    ))

    fig.add_trace(go.Scatter(
        x=data.index,
        y=data["Influenza_Cases"],
        mode="lines",
        name="Influenza",
        line=dict(width=3)
    ))

    fig.add_trace(go.Scatter(
        x=data.index,
        y=data["Respiratory_Infections"],
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


# -----------------------------
# 6️⃣ Interactive Pie Chart
# -----------------------------
def generate_distribution_pie(data):

    totals = [
        data["Malaria_Cases"].sum(),
        data["Influenza_Cases"].sum(),
        data["Respiratory_Infections"].sum()
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
