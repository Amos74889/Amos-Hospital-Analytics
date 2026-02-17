from flask import Flask, render_template, request
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from models import load_data, train_random_forest, train_arima
from config import DevelopmentConfig  # <-- Load config

# Flask app
app = Flask(__name__)
app.config.from_object(DevelopmentConfig)  # <-- Apply config

# Ensure plots folder exists
plots_folder = os.path.join(app.static_folder, "plots")
os.makedirs(plots_folder, exist_ok=True)

# Load dataset
data = load_data()

def plot_disease_trends(df):
    plt.figure(figsize=(10,6))
    sns.lineplot(x='Date', y='Malaria_Cases', data=df, label='Malaria')
    sns.lineplot(x='Date', y='Influenza_Cases', data=df, label='Influenza')
    sns.lineplot(x='Date', y='Respiratory_Infections', data=df, label='Respiratory Infections')
    plt.title("Disease Trends in Kenya")
    plt.xlabel("Date")
    plt.ylabel("Cases")
    plt.xticks(rotation=45)

    plot_path = os.path.join("plots", "disease_trends.png")  # Relative to static folder
    plt.tight_layout()
    plt.savefig(os.path.join(app.static_folder, plot_path))
    plt.close()
    return plot_path

@app.route("/", methods=["GET", "POST"])
def index():
    results = {}
    plot_path = plot_disease_trends(data)

    if request.method == "POST":
        file = request.files.get("file")
        if file:
            data_uploaded = pd.read_csv(file)
        else:
            data_uploaded = data

        # Train models
        rf_result = train_random_forest(data_uploaded)
        arima_model = train_arima(data_uploaded)

        results = {
            "rf_mae": rf_result.get("MAE", "N/A"),
            "rf_rmse": rf_result.get("RMSE", "N/A"),
            "arima_summary": arima_model.summary().as_text()
        }

    return render_template("index.html", results=results, plot_path=plot_path)

if __name__ == "__main__":
    app.run(debug=app.config["DEBUG"])
