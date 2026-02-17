import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_absolute_error, mean_squared_error
from statsmodels.tsa.arima.model import ARIMA
import matplotlib.pyplot as plt
import seaborn as sns
import pickle
import os

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
    # Create a synthetic hospital admission target
    data["Hospital_Admissions"] = (
        0.3 * data["Influenza_Cases"] +
        0.4 * data["Malaria_Cases"] +
        0.3 * data["Respiratory_Infections"]
    )

    # Add time-based features
    data["Month"] = data.index.month
    data["Year"] = data.index.year

    return data

# -----------------------------
# 3️⃣ Random Forest Model
# -----------------------------
def train_random_forest(data):
    features = ["Influenza_Cases", "Malaria_Cases", "Respiratory_Infections", "Month", "Year"]
    X = data[features]
    y = data["Hospital_Admissions"]

    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, shuffle=False)

    model = RandomForestRegressor(n_estimators=200, random_state=42)
    model.fit(X_train, y_train)

    predictions = model.predict(X_test)
    mae = mean_absolute_error(y_test, predictions)
    rmse = np.sqrt(mean_squared_error(y_test, predictions))

    print("Random Forest Performance:")
    print("MAE:", mae)
    print("RMSE:", rmse)

    # Save model
    os.makedirs("models", exist_ok=True)
    with open("models/random_forest_model.pkl", "wb") as f:
        pickle.dump(model, f)

    return {"model": model, "MAE": mae, "RMSE": rmse}

# -----------------------------
# 4️⃣ ARIMA Model
# -----------------------------
def train_arima(data):
    series = data["Hospital_Admissions"]
    model = ARIMA(series, order=(2,1,2))
    model_fit = model.fit()
    print("ARIMA Model Trained")
    return model_fit

# -----------------------------
# 5️⃣ Plot Disease Trends (Blue Theme)
# -----------------------------
def plot_disease_trends(data):
    plt.figure(figsize=(10,6))
    sns.set_style("whitegrid")
    sns.set_palette(["#1E90FF", "#87CEFA", "#ADD8E6"])  # Blue shades

    sns.lineplot(x="Date", y="Malaria_Cases", data=data, label="Malaria", linewidth=2.5)
    sns.lineplot(x="Date", y="Influenza_Cases", data=data, label="Influenza", linewidth=2.5)
    sns.lineplot(x="Date", y="Respiratory_Infections", data=data, label="Respiratory Infections", linewidth=2.5)

    plt.title("Disease Trends in Kenya", fontsize=16, color="#1E90FF")
    plt.xlabel("Date")
    plt.ylabel("Cases")
    plt.xticks(rotation=45)
    plt.legend(title="Disease")

    os.makedirs("static/plots", exist_ok=True)
    plot_path = "static/plots/disease_trends.png"
    plt.tight_layout()
    plt.savefig(plot_path)
    plt.close()
    return plot_path

# -----------------------------
# 6️⃣ Main Execution
# -----------------------------
if __name__ == "__main__":
    data = load_data()
    data = prepare_features(data)

    print("Dataset Loaded Successfully")
    print(data.head())

    rf_result = train_random_forest(data)
    arima_model = train_arima(data)
    plot_path = plot_disease_trends(data)

    print("Models trained successfully!")
    print(f"Disease plot saved at: {plot_path}")
