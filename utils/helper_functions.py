import pandas as pd
import pickle
import os

def load_data(filepath):
    """Load dataset from the specified filepath."""
    return pd.read_csv(filepath)

def load_model(model_path):
    """Load the trained machine learning pipeline from a pickle file."""
    if os.path.exists(model_path):
        with open(model_path, 'rb') as f:
            return pickle.load(f)
    return None

def prepare_input_data(
    country, region, year_normalized, vehicle_segment, petrol_car_sales, diesel_car_sales, 
    charging_stations, fast_chargers_share, average_ev_range,
    fuel_price, electricity_price, fuel_to_electric_ratio, economic_index, policy_index,
    policy_index_lagged_1y, environmental_stringency_ratio
):
    """
    Format raw Streamlit web inputs into a standard Pandas DataFrame matching the exact 
    structure required by the strictly refined scikit-learn preprocessing pipeline.
    """
    data = {
        'country': [country],
        'region': [region],
        'year_normalized': [year_normalized],
        'vehicle_segment': [vehicle_segment],
        'petrol_car_sales(units)': [petrol_car_sales],
        'diesel_car_sales(units)': [diesel_car_sales],
        'charging_stations(units)': [charging_stations],
        'fast_chargers_share(%)': [fast_chargers_share],
        'average_ev_range(km)': [average_ev_range],
        'fuel_price(usd/ltr)': [fuel_price],
        'electricity_price(usd/kwh)': [electricity_price],
        'fuel_to_electric_ratio': [fuel_to_electric_ratio],
        'economic_index': [economic_index],
        'policy_index': [policy_index],
        'policy_index_lagged_1y': [policy_index_lagged_1y],
        'environmental_stringency_ratio': [environmental_stringency_ratio]
    }
    
    return pd.DataFrame(data)
