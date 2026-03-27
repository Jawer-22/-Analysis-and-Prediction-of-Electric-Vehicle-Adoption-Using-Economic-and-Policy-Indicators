# Analysis & Prediction of EV Adoption Using Economic and Policy Indicators

This project provides a comprehensive end-to-end Machine Learning pipeline to analyze and predict the market share of Electric Vehicles (EVs) across different regions and country profiles, driven by economic indicators and policy stringency.

## Project Overview
The goal of this analysis is to identify the primary drivers of EV adoption—such as fuel prices, government subsidies, and infrastructure development—and build a predictive model to estimate future market share.

## Data Science Pipeline
1.  **Data Understanding**: Initial exploration of global EV sales and macroeconomic datasets.
2.  **Data Cleaning**: Standardizing units, handling missing values, and ensuring data integrity.
3.  **Feature Engineering**: Creation of high-impact features like `economic_index`, `policy_index`, and `average_ev_range`.
4.  **Exploratory Data Analysis (EDA)**: Visualizing correlations between policy stringency and market penetration.
5.  **Model Building**: Training and tuning **Random Forest** and **XGBoost** regressors.
6.  **Model Evaluation**: Comparing performance using RMSE, MAE, and R² scores. (XGBoost identified as the champion model).
7.  **Model Explainability**: Interpreting feature importance (Gini importance) to understand global adoption drivers.

## Directory Structure
- `Data/`: Datasets at various stages (Raw, Cleaned, Featured).
- `Notebooks/`: Sequential Jupyter notebooks from 01 to 07.
- `Models/`: Serialized models including the final `trained_model.pkl`.
- `App/`: Source code for the web application and prediction interface.
- `utils/`: Core helper functions for model loading and data preparation.

## Deployment
A web-based prediction interface is available to interact with the model. This allows users to input economic and policy scenarios and receive real-time EV market share predictions.


