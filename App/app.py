import streamlit as st
import pandas as pd
import numpy as np
import pickle
import sys
import os

# Ensure utils can be imported properly
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from utils.helper_functions import load_model, prepare_input_data

st.set_page_config(page_title="EV Market Predictor", layout="wide", page_icon="⚡")

st.title("Electric Vehicle Market Share Predictor ⚡🚗")
st.markdown("Enter the economic, infrastructural, and policy data for a region below to predict the resultant **EV Market Share** using our advanced XGBoost machine learning model.")
st.markdown("---")

# Provide inputs
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("Categorical Info")
    country = st.text_input("Country Name", value="Norway")
    region = st.selectbox("Region", ["Europe", "North America", "APAC", "Oceania", "South America", "Africa"], index=0)
    vehicle_segment = st.selectbox("Vehicle Segment", ["mass_market", "premium", "commercial"], index=0)

with col2:
    st.subheader("Policy & Economics")
    policy_index = st.slider("Policy Index (Subsidies & Regulations)", 0, 8000, 2000)
    policy_index_lagged_1y = st.slider("Policy Index (Lagged 1 Year)", 0, 8000, 1500)
    economic_index = st.number_input("Economic Index (GDP * Urban%)", value=5000.0)
    environmental_stringency_ratio = st.number_input("Environmental Stringency Ratio", value=2.0)

with col3:
    st.subheader("Infrastructure & Range")
    charging_stations = st.number_input("Charging Stations (Units)", value=500)
    fast_chargers_share = st.slider("Fast Chargers Share (%)", 0.0, 100.0, 10.0)
    average_ev_range = st.number_input("Average EV Range (km)", value=300)

st.markdown("---")
st.subheader("Cost & General Data")
col4, col5, col6, col7 = st.columns(4)
fuel_price = col4.number_input("Fuel Price (USD/ltr)", value=1.5)
electricity_price = col5.number_input("Electricity Price (USD/kWh)", value=0.15)
fuel_to_electric_ratio = col6.number_input("Fuel to Electric Ratio", value=10.0)
year_normalized = col7.number_input("Years since 2010", value=10)

# Non-predictive volume baselines
petrol_car_sales = 10000 
diesel_car_sales = 5000

st.markdown("---")

# Prediction Button
if st.button("Predict Market Share 🚀", use_container_width=True):
    model_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'Models', 'trained_model.pkl'))
    model = load_model(model_path)
    
    if model is None:
         st.error("Model not found! Please run Notebooks 05 and 06 to train and save `trained_model.pkl` to the `Models/` directory first.")
    else:
        # Prepare data
        input_data = prepare_input_data(
            country, region, year_normalized, vehicle_segment, petrol_car_sales, diesel_car_sales, 
            charging_stations, fast_chargers_share, average_ev_range,
            fuel_price, electricity_price, fuel_to_electric_ratio, economic_index, policy_index,
            policy_index_lagged_1y, environmental_stringency_ratio
        )
        
        try:
            # Predict
            prediction = model.predict(input_data)[0]
            # Assure no negative predictions
            prediction = max(0.0, min(100.0, prediction))
            
            st.success(f"### Predicted EV Market Share: {prediction:.2f}%")
            st.balloons()
        except Exception as e:
            st.error(f"Prediction Pipeline Failed: {e}\\n\\nPlease ensure all features perfectly match the exact naming and sequence of the training data.")
