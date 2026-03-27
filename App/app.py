import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import pickle
import sys
import os

# Ensure utils can be imported properly
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from utils.helper_functions import load_model, prepare_input_data

st.set_page_config(page_title="EV Market Premium Dashboard", layout="wide", page_icon="⚡")

# Mapped selection dictionary
COUNTRY_MAP = {
    "🇳🇴 Norway": "Norway", "🇸🇪 Sweden": "Sweden", "🇳🇱 Netherlands": "Netherlands", "🇦🇹 Austria": "Austria",
    "🇨🇭 Switzerland": "Switzerland", "🇬🇧 United Kingdom": "United Kingdom", "🇫🇷 France": "France",
    "🇩🇪 Germany": "Germany", "🇪🇸 Spain": "Spain", "🇵🇱 Poland": "Poland", "🇨🇦 Canada": "Canada",
    "🇺🇸 United States": "United States", "🇲🇽 Mexico": "Mexico", "🇰🇷 South Korea": "South Korea",
    "🇹🇭 Thailand": "Thailand", "🇮🇳 India": "India", "🇦🇺 Australia": "Australia"
}

# --- SIDEBAR CONFIGURATION ---
st.sidebar.title("🌍 Prediction Engine")
st.sidebar.markdown("Configure the regional ecosystem variables below.")

country_display = st.sidebar.selectbox("Target Market Selection", list(COUNTRY_MAP.keys()), index=0)
country = COUNTRY_MAP[country_display]
region = st.sidebar.selectbox("Continental Region", ["Europe", "North America", "APAC", "Oceania", "South America", "Africa"], index=0)
vehicle_segment = st.sidebar.selectbox("Vehicle Segment Strategy", ["mass_market", "premium", "commercial"], index=0)

st.sidebar.markdown("### Economics & Cost Dynamics")
economic_index = st.sidebar.number_input("Economic Index (GDP * Urban%)", value=5000.0)
fuel_price = st.sidebar.number_input("Internal Combustion Fuel Price (USD/ltr)", value=1.5)
electricity_price = st.sidebar.number_input("Electricity Cost (USD/kWh)", value=0.15)
fuel_to_electric_ratio = st.sidebar.number_input("Fuel-to-Electric Price Ratio", value=10.0)

st.sidebar.markdown("### Structural Infrastructure & Policy")
policy_index = st.sidebar.slider("Current Policy Index (Subsidies/Regulations)", 0, 8000, 2000)
policy_index_lagged_1y = st.sidebar.slider("Lagged 1-Year Policy Index", 0, 8000, 1500)
environmental_stringency_ratio = st.sidebar.number_input("Environmental Stringency Factor", value=2.0)
charging_stations = st.sidebar.number_input("Public Charging Stations (Units)", value=500)
fast_chargers_share = st.sidebar.slider("Fast Charging Capacity Share (%)", 0.0, 100.0, 10.0)
average_ev_range = st.sidebar.number_input("Average Target EV Range (km)", value=300)
year_normalized = st.sidebar.number_input("Years Progression since 2010", value=10)

# Non-predictive volume baselines locked historically
petrol_car_sales = 10000 
diesel_car_sales = 5000

# --- MAIN DASHBOARD VIEW ---
st.title("Electric Vehicle Market Prediction Engine ⚡🚗")
st.markdown("This premium dashboard autonomously evaluates macroeconomic shifts, infrastructure expansion, and political subsidies to robustly forecast the resulting adoption integrations of Electric Vehicles.")

# Metrics Cards
c1, c2, c3, c4 = st.columns(4)
c1.metric("Assigned Target Market", country_display)
c2.metric("Economic Baseline", f"{economic_index:,.0f}")
c3.metric("Policy Velocity", f"{policy_index:,.0f}")
c4.metric("Infrastructural Range", f"{average_ev_range} km")

st.markdown("---")

# Execution Space
predict_clicked = st.button("Generate Market Share Prediction 🚀", use_container_width=True)

if predict_clicked:
    model_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'Models', 'trained_model.pkl'))
    model = load_model(model_path)
    
    if model is None:
         st.error("Algorithm missing! Please ensure `trained_model.pkl` is safely deployed in the `Models/` root directory.")
    else:
        # Preprocessing translation layer
        input_data = prepare_input_data(
            country, region, year_normalized, vehicle_segment, petrol_car_sales, diesel_car_sales, 
            charging_stations, fast_chargers_share, average_ev_range,
            fuel_price, electricity_price, fuel_to_electric_ratio, economic_index, policy_index,
            policy_index_lagged_1y, environmental_stringency_ratio
        )
        
        try:
            # Extraction
            prediction = model.predict(input_data)[0]
            prediction = max(0.0, min(100.0, prediction))
            
            st.success(f"### Algorithm Successful: Computed an EV Market Penetration Rate of {prediction:.2f}%")
            
            # Premium Visual UI Update
            fig = go.Figure(go.Indicator(
                mode = "gauge+number",
                value = prediction,
                title = {'text': "Live Estimated Adoption Funnel", 'font': {'size': 24}},
                number = {'suffix': "%", 'font': {'size': 70, 'color': '#02804c'}},
                gauge = {
                    'axis': {'range': [None, 100], 'tickwidth': 2, 'tickcolor': "darkgray"},
                    'bar': {'color': "#00c968", 'thickness': 0.25},
                    'bgcolor': "white",
                    'steps': [
                        {'range': [0, 15], 'color': '#ff4b4b'},
                        {'range': [15, 35], 'color': '#ffb020'},
                        {'range': [35, 65], 'color': '#bcf069'},
                        {'range': [65, 100], 'color': '#00d500'}],
                    'threshold': {
                        'line': {'color': "black", 'width': 4},
                        'thickness': 0.75,
                        'value': prediction
                    }
                }
            ))
            fig.update_layout(height=400, margin=dict(l=10, r=10, t=50, b=10))
            st.plotly_chart(fig, use_container_width=True)
            
            if prediction > 15: st.balloons()
            
        except Exception as e:
            st.error(f"Computation Pipeline Crash: {e}\\n\\nPlease verify variable parity matches expected XGBoost parameter thresholds.")
