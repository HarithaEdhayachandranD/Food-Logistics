"""
QuickBite Delivery Services – Synthetic Dataset Generator
Round 2: EDA and Feature Engineering
Generates: orders.csv, riders.csv, routes.csv, weather.csv, delivery.csv
with intentional data quality issues as specified in the brief.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

np.random.seed(42)
random.seed(42)

# ─── Constants ───────────────────────────────────────────────────────────────
N_ORDERS   = 1200
N_RIDERS   = 80
N_ZONES    = 8

ZONES_CLEAN = ['North', 'South', 'East', 'West', 'Central', 'Airport', 'Suburbs', 'Downtown']
# Inconsistent zone name variations injected for data quality issues
ZONES_MESSY = ZONES_CLEAN + [
    'north', 'SOUTH', 'East Zone', 'west', 'central', 'AIRPORT', 'suburbs', 'Down Town',
    'Nrth', 'Sth'
]

RESTAURANTS = [
    'Spice Garden', 'Pizza Palace', 'Burger Barn', 'Noodle House',
    'Curry Corner', 'Sushi Spot', 'Taco Town', 'Salad Station',
    'Wrap World', 'BBQ Brothers', 'Dim Sum Den', 'Pasta Park'
]

SHIFTS = ['Morning', 'Afternoon', 'Evening', 'Night']
EXPERIENCE_LEVELS = ['Junior', 'Mid', 'Senior']
TRAFFIC_LEVELS = ['Low', 'Medium', 'High', 'Very High']
WEATHER_CONDITIONS = ['Clear', 'Cloudy', 'Light Rain', 'Heavy Rain', 'Fog']
STATUSES = ['Delivered', 'Delivered', 'Delivered', 'Late', 'Late', 'Failed']  # weighted

start_date = datetime(2024, 1, 1)
end_date   = datetime(2024, 6, 30)

def random_datetime(start, end):
    delta = end - start
    random_seconds = random.randint(0, int(delta.total_seconds()))
    return start + timedelta(seconds=random_seconds)

# ─── 1. orders.csv ────────────────────────────────────────────────────────────
print("Generating orders.csv ...")
order_ids = [f"ORD{str(i).zfill(5)}" for i in range(1, N_ORDERS + 1)]

# Inject ~3% duplicate order IDs
dup_indices = np.random.choice(range(N_ORDERS), size=35, replace=False)
for idx in dup_indices:
    order_ids[idx] = order_ids[idx - 1]

order_times = [random_datetime(start_date, end_date) for _ in range(N_ORDERS)]
# Inject ~5% missing timestamps
missing_time_idx = np.random.choice(range(N_ORDERS), size=60, replace=False)
order_times_series = pd.Series(order_times)
order_times_series.iloc[missing_time_idx] = pd.NaT

zones_orders = [random.choice(ZONES_MESSY) for _ in range(N_ORDERS)]
restaurants  = [random.choice(RESTAURANTS) for _ in range(N_ORDERS)]

# Order value: base value influenced by restaurant & some noise
base_values = {r: np.random.uniform(120, 600) for r in RESTAURANTS}
order_values = [round(base_values[r] + np.random.normal(0, 60), 2) for r in restaurants]
# Inject a few negative/zero values
neg_idx = np.random.choice(range(N_ORDERS), size=8, replace=False)
for idx in neg_idx:
    order_values[idx] = round(np.random.uniform(-50, 0), 2)

# Inject some missing values
missing_val_idx = np.random.choice(range(N_ORDERS), size=20, replace=False)
order_values_series = pd.Series(order_values, dtype=float)
order_values_series.iloc[missing_val_idx] = np.nan

orders_df = pd.DataFrame({
    'order_id'       : order_ids,
    'order_time'     : order_times_series,
    'restaurant'     : restaurants,
    'customer_zone'  : zones_orders,
    'order_value_inr': order_values_series
})
orders_df.to_csv('/Users/navins/Documents/Hari_DS_Spring/orders.csv', index=False)
print(f"  orders.csv: {orders_df.shape}")

# ─── 2. riders.csv ────────────────────────────────────────────────────────────
print("Generating riders.csv ...")
rider_ids = [f"RDR{str(i).zfill(3)}" for i in range(1, N_RIDERS + 1)]
experience = [random.choice(EXPERIENCE_LEVELS) for _ in range(N_RIDERS)]
shifts     = [random.choice(SHIFTS) for _ in range(N_RIDERS)]

# Years of experience (correlated with level)
exp_years = []
for e in experience:
    if e == 'Junior':  exp_years.append(round(np.random.uniform(0, 1.5), 1))
    elif e == 'Mid':   exp_years.append(round(np.random.uniform(1.5, 4), 1))
    else:              exp_years.append(round(np.random.uniform(4, 10), 1))

ratings = [round(np.clip(np.random.normal(3.8, 0.6), 1, 5), 1) for _ in range(N_RIDERS)]

# Inject some missing ratings
miss_rating = np.random.choice(range(N_RIDERS), size=8, replace=False)
ratings_series = pd.Series(ratings, dtype=float)
ratings_series.iloc[miss_rating] = np.nan

riders_df = pd.DataFrame({
    'rider_id'        : rider_ids,
    'experience_level': experience,
    'experience_years': exp_years,
    'shift'           : shifts,
    'avg_rating'      : ratings_series
})
riders_df.to_csv('/Users/navins/Documents/Hari_DS_Spring/riders.csv', index=False)
print(f"  riders.csv: {riders_df.shape}")

# ─── 3. routes.csv ────────────────────────────────────────────────────────────
print("Generating routes.csv ...")
route_ids    = [f"RTE{str(i).zfill(5)}" for i in range(1, N_ORDERS + 1)]
zones_routes = [random.choice(ZONES_MESSY) for _ in range(N_ORDERS)]
distances_km = [round(np.random.uniform(0.8, 22), 2) for _ in range(N_ORDERS)]
traffic      = [random.choice(TRAFFIC_LEVELS) for _ in range(N_ORDERS)]

# Traffic multiplier for realism
traffic_mult = {'Low': 1.0, 'Medium': 1.3, 'High': 1.7, 'Very High': 2.3}

# Inject some missing distances
miss_dist = np.random.choice(range(N_ORDERS), size=25, replace=False)
distances_series = pd.Series(distances_km, dtype=float)
distances_series.iloc[miss_dist] = np.nan

routes_df = pd.DataFrame({
    'route_id'    : route_ids,
    'zone'        : zones_routes,
    'distance_km' : distances_series,
    'traffic_level': traffic
})
routes_df.to_csv('/Users/navins/Documents/Hari_DS_Spring/routes.csv', index=False)
print(f"  routes.csv: {routes_df.shape}")

# ─── 4. weather.csv ───────────────────────────────────────────────────────────
print("Generating weather.csv ...")
# One weather record per hour for the date range
weather_times = pd.date_range(start=start_date, end=end_date, freq='h')
n_weather = len(weather_times)

# Temperature follows a sinusoidal daily pattern
hours = weather_times.hour
temperatures = 22 + 8 * np.sin((hours - 6) * np.pi / 12) + np.random.normal(0, 2, n_weather)

# Rain flag and intensity
rain_flag = np.random.choice([0, 1], size=n_weather, p=[0.7, 0.3])
rainfall_mm = np.where(rain_flag == 1, np.random.exponential(5, n_weather), 0).round(1)

visibility_km = np.clip(10 - rainfall_mm * 0.8 + np.random.normal(0, 1, n_weather), 0.5, 10).round(1)

weather_conditions_list = []
for rf, rm in zip(rain_flag, rainfall_mm):
    if rf == 0:
        weather_conditions_list.append(random.choice(['Clear', 'Cloudy']))
    elif rm < 5:
        weather_conditions_list.append('Light Rain')
    elif rm < 15:
        weather_conditions_list.append('Heavy Rain')
    else:
        weather_conditions_list.append('Heavy Rain')

# Inject ~5% missing weather readings
miss_weather_idx = np.random.choice(range(n_weather), size=int(n_weather * 0.05), replace=False)
temp_series  = pd.Series(temperatures.round(1))
rain_series  = pd.Series(rainfall_mm)
vis_series   = pd.Series(visibility_km)
temp_series.iloc[miss_weather_idx[:len(miss_weather_idx)//3]]  = np.nan
rain_series.iloc[miss_weather_idx[len(miss_weather_idx)//3:2*len(miss_weather_idx)//3]] = np.nan
vis_series.iloc[miss_weather_idx[2*len(miss_weather_idx)//3:]] = np.nan

weather_df = pd.DataFrame({
    'datetime'          : weather_times,
    'temperature_c'     : temp_series,
    'rain_flag'         : rain_series,
    'rainfall_mm'       : rain_series,
    'visibility_km'     : vis_series,
    'weather_condition' : weather_conditions_list
})
weather_df.to_csv('/Users/navins/Documents/Hari_DS_Spring/weather.csv', index=False)
print(f"  weather.csv: {weather_df.shape}")

# ─── 5. delivery.csv ──────────────────────────────────────────────────────────
print("Generating delivery.csv ...")

# Assign rider IDs to orders
assigned_riders = [random.choice(rider_ids) for _ in range(N_ORDERS)]

# Pickup time = order_time + pickup_delay (influenced by restaurant & traffic)
# restaurant_pickup_base (minutes)
rest_pickup_base = {r: np.random.uniform(5, 25) for r in RESTAURANTS}

delivery_records = []
for i in range(N_ORDERS):
    oid  = order_ids[i]
    ot   = order_times_series.iloc[i]
    rest = restaurants[i]
    traf = traffic[i]
    dist = distances_series.iloc[i] if not pd.isna(distances_series.iloc[i]) else np.random.uniform(3, 12)

    rider_id = assigned_riders[i]
    exp_lvl  = riders_df[riders_df['rider_id'] == rider_id]['experience_level'].values[0]

    # Base pickup delay (minutes)
    pickup_delay = rest_pickup_base[rest] + np.random.normal(0, 5)
    pickup_delay = max(pickup_delay, 2)

    # Traffic multiplier on delivery (not pickup)
    t_mult = traffic_mult.get(traf, 1.0)

    # Experience modifier
    exp_modifier = {'Junior': 1.15, 'Mid': 1.0, 'Senior': 0.88}[exp_lvl]

    # Delivery speed: base 3 min/km, modified
    delivery_speed = 3 * t_mult * exp_modifier + np.random.normal(0, 0.5)
    delivery_speed = max(delivery_speed, 1)
    delivery_duration = dist * delivery_speed

    if pd.isna(ot):
        pickup_time    = pd.NaT
        delivered_time = pd.NaT
    else:
        pickup_time    = ot + timedelta(minutes=pickup_delay)
        delivered_time = pickup_time + timedelta(minutes=delivery_duration)

    # Status
    total_time = pickup_delay + delivery_duration
    if total_time > 60:
        status = random.choice(['Late', 'Failed'])
    elif total_time > 45:
        status = 'Late' if random.random() > 0.3 else 'Delivered'
    else:
        status = 'Delivered'

    delivery_records.append({
        'delivery_id'   : f"DLV{str(i+1).zfill(5)}",
        'order_id'      : oid,
        'rider_id'      : rider_id,
        'route_id'      : route_ids[i],
        'pickup_time'   : pickup_time,
        'delivered_time': delivered_time,
        'status'        : status
    })

delivery_df = pd.DataFrame(delivery_records)

# Inject ~5% missing timestamps in delivery
miss_del_idx = np.random.choice(range(N_ORDERS), size=55, replace=False)
delivery_df.loc[miss_del_idx[:28], 'pickup_time']    = pd.NaT
delivery_df.loc[miss_del_idx[28:], 'delivered_time'] = pd.NaT

# Inject a few negative durations by swapping pickup and delivered times
neg_dur_idx = np.random.choice(range(N_ORDERS), size=10, replace=False)
for idx in neg_dur_idx:
    pt = delivery_df.loc[idx, 'pickup_time']
    dt = delivery_df.loc[idx, 'delivered_time']
    if pd.notna(pt) and pd.notna(dt):
        delivery_df.loc[idx, 'pickup_time']    = dt
        delivery_df.loc[idx, 'delivered_time'] = pt

delivery_df.to_csv('/Users/navins/Documents/Hari_DS_Spring/delivery.csv', index=False)
print(f"  delivery.csv: {delivery_df.shape}")

print("\n✅ All 5 datasets generated successfully!")
print("\nSummary:")
for name, df in [('orders', orders_df), ('riders', riders_df), ('routes', routes_df),
                 ('weather', weather_df), ('delivery', delivery_df)]:
    print(f"  {name}.csv: {df.shape[0]} rows × {df.shape[1]} cols")
