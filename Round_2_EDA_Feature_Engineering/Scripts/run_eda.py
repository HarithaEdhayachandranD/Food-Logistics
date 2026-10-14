"""
QuickBite Delivery Services – Round 2: EDA & Feature Engineering
Standalone execution script (mirrors the Jupyter notebook exactly)
Run: python3 run_eda.py
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')   # headless backend – no display needed
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
import warnings
from scipy import stats
from sklearn.preprocessing import LabelEncoder
from sklearn.feature_selection import SelectKBest, f_regression
from sklearn.ensemble import RandomForestRegressor
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler

warnings.filterwarnings('ignore')
pd.set_option('display.max_columns', 50)
pd.set_option('display.float_format', '{:.2f}'.format)

# ─── Plotting theme ────────────────────────────────────────────────────────────
PALETTE = ['#E63946','#457B9D','#2A9D8F','#E9C46A','#F4A261','#264653','#A8DADC','#1D3557']
plt.rcParams.update({
    'figure.facecolor' : '#0F1117',
    'axes.facecolor'   : '#1A1D27',
    'axes.edgecolor'   : '#2E3250',
    'axes.labelcolor'  : '#C8D0E7',
    'xtick.color'      : '#8892B0',
    'ytick.color'      : '#8892B0',
    'text.color'       : '#CDD6F4',
    'grid.color'       : '#2E3250',
    'grid.linestyle'   : '--',
    'grid.alpha'       : 0.5,
    'axes.titlesize'   : 14,
    'axes.titleweight' : 'bold',
    'axes.titlepad'    : 12,
    'axes.titlecolor'  : '#CDD6F4',
    'font.family'      : 'DejaVu Sans',
    'figure.dpi'       : 130,
})

def savefig(name):
    plt.savefig(f'{name}.png', dpi=150, bbox_inches='tight',
                facecolor=plt.rcParams['figure.facecolor'])
    plt.close()
    print(f'  Saved: {name}.png')

print('='*70)
print('  QUICKBITE DELIVERY SERVICES – ROUND 2: EDA & FEATURE ENGINEERING')
print('='*70)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2: LOAD DATASETS
# ══════════════════════════════════════════════════════════════════════════════
print('\n[2] Loading datasets...')
orders   = pd.read_csv('orders.csv',   parse_dates=['order_time'])
riders   = pd.read_csv('riders.csv')
routes   = pd.read_csv('routes.csv')
weather  = pd.read_csv('weather.csv',  parse_dates=['datetime'])
delivery = pd.read_csv('delivery.csv', parse_dates=['pickup_time','delivered_time'])

datasets = {'orders':orders,'riders':riders,'routes':routes,'weather':weather,'delivery':delivery}
for name, df in datasets.items():
    print(f'  {name:10s}: {df.shape[0]:>5,} rows × {df.shape[1]} cols')

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3: DATA QUALITY ASSESSMENT
# ══════════════════════════════════════════════════════════════════════════════
print('\n[3] Data Quality Assessment')
print('\n  --- Missing Values ---')
for name, df in datasets.items():
    mv = df.isnull().sum()
    mv = mv[mv > 0]
    if not mv.empty:
        for col, cnt in mv.items():
            print(f'    {name}.{col}: {cnt} missing ({cnt/len(df)*100:.1f}%)')
    else:
        print(f'    {name}: No missing values')

print('\n  --- Duplicates ---')
print(f'    orders – duplicate order_id: {orders["order_id"].duplicated().sum()}')
print(f'    delivery – full row dups   : {delivery.duplicated().sum()}')

print('\n  --- Zone Name Inconsistencies ---')
print(f'    orders unique zones : {sorted(orders["customer_zone"].dropna().unique())}')
print(f'    routes unique zones : {sorted(routes["zone"].dropna().unique())}')

print('\n  --- Unusual Values ---')
neg_orders = orders[orders['order_value_inr'] < 0]
print(f'    Negative order values    : {len(neg_orders)}')
valid_d = delivery.dropna(subset=['pickup_time','delivered_time'])
neg_dur  = valid_d[valid_d['delivered_time'] < valid_d['pickup_time']]
print(f'    Negative durations       : {len(neg_dur)}')
print(f'    Delivery status counts   : {delivery["status"].value_counts().to_dict()}')

# Missing values heatmap
fig, axes = plt.subplots(1, 3, figsize=(18, 5))
fig.suptitle('Missing Values Heatmap by Dataset', fontsize=15, fontweight='bold', y=1.02)
for ax, (name, df_p) in zip(axes, [('Orders',orders),('Routes',routes),('Delivery',delivery)]):
    mask = df_p.isnull()
    sns.heatmap(mask, ax=ax, cbar=False, yticklabels=False, cmap=['#1A1D27','#E63946'])
    ax.set_title(f'{name}\n({mask.sum().sum()} missing)')
    ax.set_xticklabels(ax.get_xticklabels(), rotation=35, ha='right', fontsize=9)
    ax.legend(handles=[mpatches.Patch(color='#E63946',label='Missing'),
                        mpatches.Patch(color='#1A1D27',label='Present')],
              loc='upper right', fontsize=8)
plt.tight_layout()
savefig('fig01_missing_heatmap')

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4: DATA CLEANING
# ══════════════════════════════════════════════════════════════════════════════
print('\n[4] Data Cleaning')

# Step 1: Standardise zone names
ZONE_MAP = {
    'north':'North','nrth':'North','south':'South','sth':'South',
    'east':'East','east zone':'East','west':'West','central':'Central',
    'airport':'Airport','suburbs':'Suburbs','downtown':'Downtown','down town':'Downtown'
}
def standardise_zone(val):
    if pd.isna(val): return val
    return ZONE_MAP.get(val.strip().lower(), val.strip().title())

orders['customer_zone'] = orders['customer_zone'].apply(standardise_zone)
routes['zone']          = routes['zone'].apply(standardise_zone)
print(f'  Step 1: Zone names standardised → {sorted(orders["customer_zone"].dropna().unique())}')

# Step 2: Remove duplicate order IDs
before = len(orders)
orders = orders.drop_duplicates(subset='order_id', keep='first').reset_index(drop=True)
print(f'  Step 2: Removed {before-len(orders)} duplicate order_ids. Remaining: {len(orders)}')

# Step 3: Fix negative order values
orders.loc[orders['order_value_inr'] < 0, 'order_value_inr'] = np.nan
orders['order_value_inr'] = orders.groupby('restaurant')['order_value_inr'].transform(
    lambda x: x.fillna(x.median()))
orders['order_value_inr'].fillna(orders['order_value_inr'].median(), inplace=True)
print(f'  Step 3: Negative/missing order values fixed. NaN left: {orders["order_value_inr"].isnull().sum()}')

# Step 4: Fix negative durations
mask_neg = delivery['delivered_time'] < delivery['pickup_time']
delivery.loc[mask_neg, ['pickup_time','delivered_time']] = \
    delivery.loc[mask_neg, ['delivered_time','pickup_time']].values
print(f'  Step 4: Swapped {mask_neg.sum()} rows with negative delivery durations')

# Step 5: Drop missing timestamps
b_o = len(orders)
orders = orders.dropna(subset=['order_time']).reset_index(drop=True)
print(f'  Step 5a: Dropped {b_o-len(orders)} orders with missing order_time')
b_d = len(delivery)
delivery = delivery.dropna(subset=['pickup_time','delivered_time']).reset_index(drop=True)
print(f'  Step 5b: Dropped {b_d-len(delivery)} deliveries missing timestamps. Remaining: {len(delivery)}')

# Step 6: Impute missing distance
routes['distance_km'] = routes.groupby('zone')['distance_km'].transform(
    lambda x: x.fillna(x.median()))
routes['distance_km'].fillna(routes['distance_km'].median(), inplace=True)
print(f'  Step 6: Distance imputed. NaN left: {routes["distance_km"].isnull().sum()}')

# Step 7: Weather forward/backward fill
weather = weather.sort_values('datetime').reset_index(drop=True)
weather[['temperature_c','rain_flag','rainfall_mm','visibility_km']] = \
    weather[['temperature_c','rain_flag','rainfall_mm','visibility_km']].ffill().bfill()
riders['avg_rating'].fillna(
    riders.groupby('experience_level')['avg_rating'].transform('median'), inplace=True)
print(f'  Step 7: Weather filled. Total NaN: {weather.isnull().sum().sum()} | Rider rating NaN: {riders["avg_rating"].isnull().sum()}')

print('\n  Post-Cleaning Summary:')
for name, df in [('orders',orders),('riders',riders),('routes',routes),('weather',weather),('delivery',delivery)]:
    print(f'    {name:12s}: {len(df):>5,} rows, {df.isnull().sum().sum()} NaN remaining')

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5: MERGE & COMBINE
# ══════════════════════════════════════════════════════════════════════════════
print('\n[5] Merging datasets...')
df = (delivery
      .merge(orders,  on='order_id', how='inner')
      .merge(routes,  on='route_id', how='left')
      .merge(riders,  on='rider_id', how='left'))
print(f'  After delivery+orders+routes+riders: {df.shape}')

df['order_hour_key'] = df['order_time'].dt.floor('h')
weather_h = weather.rename(columns={'datetime':'order_hour_key'})
df = df.merge(weather_h[['order_hour_key','temperature_c','rain_flag','rainfall_mm',
                           'visibility_km','weather_condition']],
              on='order_hour_key', how='left')

df['zone'] = df['customer_zone'].combine_first(df['zone'])
df.drop(columns=['customer_zone','order_hour_key'], inplace=True, errors='ignore')

for col in ['temperature_c','rain_flag','rainfall_mm','visibility_km']:
    df[col].fillna(df[col].median(), inplace=True)
df['weather_condition'].fillna('Clear', inplace=True)
print(f'  Final master dataset: {df.shape}')

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6: DESCRIPTIVE ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
print('\n[6] Computing timing metrics...')
df['delivery_duration_min'] = ((df['delivered_time']-df['pickup_time']).dt.total_seconds()/60).round(2)
df['pickup_wait_min']       = ((df['pickup_time']-df['order_time']).dt.total_seconds()/60).round(2)
df['total_time_min']        = df['pickup_wait_min'] + df['delivery_duration_min']
df = df[(df['delivery_duration_min']>0) & (df['pickup_wait_min']>0)].reset_index(drop=True)
print(f'  Valid records: {len(df)}')

key_cols = ['delivery_duration_min','pickup_wait_min','total_time_min',
            'distance_km','order_value_inr','temperature_c','rainfall_mm','visibility_km']
print('\n  Descriptive Statistics:')
print(df[key_cols].describe().round(2).to_string())

print('\n  Categorical Distributions:')
for col in ['status','traffic_level','experience_level','shift','weather_condition','zone']:
    print(f'\n  {col}:')
    print(df[col].value_counts(dropna=False).to_string())

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7: VISUALIZATIONS
# ══════════════════════════════════════════════════════════════════════════════
print('\n[7] Creating visualizations...')

# Fig 2: Delivery time distribution
fig, axes = plt.subplots(1, 2, figsize=(16, 5))
fig.suptitle('Delivery Time Distribution', fontsize=15, fontweight='bold')
axes[0].hist(df['total_time_min'], bins=50, color='#457B9D', edgecolor='#0F1117', alpha=0.85)
axes[0].axvline(df['total_time_min'].mean(),   color='#E63946', lw=2, ls='--', label=f'Mean={df["total_time_min"].mean():.1f}')
axes[0].axvline(df['total_time_min'].median(), color='#2A9D8F', lw=2, ls='-.', label=f'Median={df["total_time_min"].median():.1f}')
axes[0].set_xlabel('Total Time (min)'); axes[0].set_ylabel('Frequency')
axes[0].set_title('Histogram'); axes[0].legend(); axes[0].grid(True)

status_order = df.groupby('status')['total_time_min'].median().sort_values(ascending=False).index
data_s = [df[df['status']==s]['total_time_min'].dropna().values for s in status_order]
bp = axes[1].boxplot(data_s, patch_artist=True, notch=True, medianprops=dict(color='#E9C46A',lw=2))
status_colors = ['#E63946','#E9C46A','#2A9D8F']
for patch, c in zip(bp['boxes'], status_colors): patch.set_facecolor(c); patch.set_alpha(0.7)
axes[1].set_xticklabels(status_order); axes[1].set_ylabel('Total Time (min)')
axes[1].set_title('Box Plot by Status'); axes[1].grid(True)
plt.tight_layout(); savefig('fig02_delivery_time_dist')

# Fig 3: Status Distribution
fig, axes = plt.subplots(1, 2, figsize=(15, 5))
fig.suptitle('Delivery Status Overview', fontsize=15, fontweight='bold')
sc = df['status'].value_counts()
c_pie = ['#2A9D8F','#E9C46A','#E63946']
wedges, texts, ats = axes[0].pie(sc.values, labels=sc.index, autopct='%1.1f%%',
                                  colors=c_pie, startangle=140,
                                  wedgeprops=dict(edgecolor='#0F1117',linewidth=2), pctdistance=0.82)
for at in ats: at.set_color('#0F1117'); at.set_fontsize(10)
axes[0].set_title('Proportion by Status')
bars = axes[1].bar(sc.index, sc.values, color=c_pie, edgecolor='#0F1117', alpha=0.85)
for bar in bars:
    axes[1].text(bar.get_x()+bar.get_width()/2, bar.get_height()+2, f'{int(bar.get_height())}',
                 ha='center', fontsize=11)
axes[1].set_xlabel('Status'); axes[1].set_ylabel('Count')
axes[1].set_title('Count by Status'); axes[1].grid(True, axis='y')
plt.tight_layout(); savefig('fig03_status_distribution')

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 8: USE-CASE ANALYTICAL QUESTIONS
# ══════════════════════════════════════════════════════════════════════════════
print('\n[8] Answering Use-Case Questions...')

# Q1: Zone delivery times
print('\n  Q1: Zone Delivery Times')
zone_time = (df.groupby('zone')['total_time_min']
               .agg(['mean','median','count'])
               .sort_values('mean', ascending=False)
               .reset_index())
zone_time.columns = ['Zone','Mean Time','Median Time','Orders']
print(zone_time.to_string(index=False))

fig, axes = plt.subplots(1, 2, figsize=(16, 5))
fig.suptitle('Q1: Delivery Time by Zone', fontsize=15, fontweight='bold')
c_z = [PALETTE[i % len(PALETTE)] for i in range(len(zone_time))]
axes[0].barh(zone_time['Zone'], zone_time['Mean Time'], color=c_z, edgecolor='#0F1117', alpha=0.85)
for bar, val in zip(axes[0].patches, zone_time['Mean Time']):
    axes[0].text(val+0.3, bar.get_y()+bar.get_height()/2, f'{val:.1f}', va='center', fontsize=10)
axes[0].set_xlabel('Mean Total Time (min)'); axes[0].set_title('Mean Delivery Time by Zone')
axes[0].grid(True, axis='x')
data_z = [df[df['zone']==z]['total_time_min'].dropna().values for z in zone_time['Zone']]
bp2 = axes[1].boxplot(data_z, patch_artist=True, vert=False, medianprops=dict(color='#E9C46A',lw=2))
for patch, c in zip(bp2['boxes'], c_z): patch.set_facecolor(c); patch.set_alpha(0.6)
axes[1].set_yticklabels(zone_time['Zone'])
axes[1].set_xlabel('Total Time (min)'); axes[1].set_title('Distribution by Zone')
axes[1].grid(True, axis='x')
plt.tight_layout(); savefig('fig04_zone_delivery_time')

# Q2: Traffic impact
print('\n  Q2: Traffic Impact')
traffic_order = ['Low','Medium','High','Very High']
t_colors      = ['#2A9D8F','#E9C46A','#F4A261','#E63946']
traffic_time  = df.groupby('traffic_level')['total_time_min'].agg(['mean','median','std','count']).reindex(traffic_order)
print(traffic_time.round(2).to_string())

fig, axes = plt.subplots(1, 2, figsize=(16, 5))
fig.suptitle('Q2: Traffic Level vs Delivery Time', fontsize=15, fontweight='bold')
axes[0].bar(traffic_order, traffic_time['mean'], color=t_colors, edgecolor='#0F1117', alpha=0.85)
axes[0].errorbar(traffic_order, traffic_time['mean'], yerr=traffic_time['std'],
                 fmt='none', color='white', capsize=5, lw=1.5)
for i, m in enumerate(traffic_time['mean']):
    axes[0].text(i, m+0.5, f'{m:.1f}', ha='center', fontsize=11)
axes[0].set_ylabel('Mean Total Time (min)'); axes[0].set_title('Delivery Time vs Traffic')
axes[0].grid(True, axis='y')
data_t = [df[df['traffic_level']==t]['total_time_min'].dropna().values for t in traffic_order]
parts = axes[1].violinplot(data_t, positions=range(len(traffic_order)), showmedians=True)
for pc, col in zip(parts['bodies'], t_colors): pc.set_facecolor(col); pc.set_alpha(0.7)
parts['cmedians'].set_color('#E9C46A')
axes[1].set_xticks(range(len(traffic_order))); axes[1].set_xticklabels(traffic_order)
axes[1].set_ylabel('Total Time (min)'); axes[1].set_title('Violin Plot by Traffic')
axes[1].grid(True, axis='y')
plt.tight_layout(); savefig('fig05_traffic_impact')

# Q3: Rain impact
print('\n  Q3: Rain Impact')
df['rain_category'] = pd.cut(df['rainfall_mm'], bins=[-np.inf,0,5,15,np.inf],
                               labels=['No Rain','Light Rain','Moderate Rain','Heavy Rain'])
rain_labels = ['No Rain','Light Rain','Moderate Rain','Heavy Rain']
rain_colors = ['#1D3557','#457B9D','#A8DADC','#E63946']
rain_means  = [df[df['rain_category']==r]['total_time_min'].mean() for r in rain_labels]

fig, axes = plt.subplots(1, 2, figsize=(16, 5))
fig.suptitle('Q3: Rain vs Delivery Delay', fontsize=15, fontweight='bold')
axes[0].bar(rain_labels, rain_means, color=rain_colors, edgecolor='#0F1117', alpha=0.9)
for i, m in enumerate(rain_means):
    axes[0].text(i, m+0.4, f'{m:.1f}', ha='center', fontsize=11)
axes[0].set_ylabel('Mean Total Time (min)'); axes[0].set_title('Delivery Time by Rain Level')
axes[0].grid(True, axis='y')
sc_plot = axes[1].scatter(df['rainfall_mm'], df['total_time_min'], c=df['rainfall_mm'],
                     cmap='RdYlGn_r', alpha=0.4, s=15)
z_val = np.polyfit(df['rainfall_mm'], df['total_time_min'], 1)
p_val_poly = np.poly1d(z_val)
xl = np.linspace(0, df['rainfall_mm'].max(), 100)
axes[1].plot(xl, p_val_poly(xl), color='#E9C46A', lw=2.5, label='Trend')
plt.colorbar(sc_plot, ax=axes[1], label='Rainfall (mm)')
axes[1].set_xlabel('Rainfall (mm)'); axes[1].set_ylabel('Total Time (min)')
axes[1].set_title('Scatter + Trend'); axes[1].legend(); axes[1].grid(True)
plt.tight_layout(); savefig('fig06_rain_impact')

no_rain  = df[df['rain_flag'] < 0.5]['total_time_min'].dropna()
yes_rain = df[df['rain_flag'] >= 0.5]['total_time_min'].dropna()
t_stat, p_ttest = stats.ttest_ind(no_rain, yes_rain)
print(f'  T-Test: t={t_stat:.3f}, p={p_ttest:.4f} → ' +
      ('Significant increase with rain' if p_ttest < 0.05 else 'No significant difference'))

# Q4: Restaurant pickup times
print('\n  Q4: Restaurant Pickup Times')
rest_pickup = (df.groupby('restaurant')['pickup_wait_min']
               .agg(['mean','median','count'])
               .sort_values('mean', ascending=False)
               .reset_index())
rest_pickup.columns = ['Restaurant','Mean Wait','Median Wait','Orders']
print(rest_pickup.to_string(index=False))

fig, axes = plt.subplots(1, 2, figsize=(18, 5))
fig.suptitle('Q4: Restaurant Pickup Wait Times', fontsize=15, fontweight='bold')
c_r = [PALETTE[i % len(PALETTE)] for i in range(len(rest_pickup))]
axes[0].barh(rest_pickup['Restaurant'], rest_pickup['Mean Wait'], color=c_r, edgecolor='#0F1117', alpha=0.85)
for bar, val in zip(axes[0].patches, rest_pickup['Mean Wait']):
    axes[0].text(val+0.1, bar.get_y()+bar.get_height()/2, f'{val:.1f}', va='center', fontsize=9)
axes[0].set_xlabel('Mean Pickup Wait (min)'); axes[0].set_title('Pickup Wait by Restaurant')
axes[0].grid(True, axis='x')
top6 = rest_pickup['Restaurant'].head(6).tolist()
d_r  = [df[df['restaurant']==r]['pickup_wait_min'].dropna().values for r in top6]
bp3  = axes[1].boxplot(d_r, patch_artist=True, medianprops=dict(color='#E9C46A',lw=2))
for patch, c in zip(bp3['boxes'], c_r[:6]): patch.set_facecolor(c); patch.set_alpha(0.7)
axes[1].set_xticklabels(top6, rotation=30, ha='right')
axes[1].set_ylabel('Pickup Wait (min)'); axes[1].set_title('Top 6 Slowest Restaurants')
axes[1].grid(True, axis='y')
plt.tight_layout(); savefig('fig07_restaurant_pickup')

# Q5: Distance vs time
print('\n  Q5: Distance vs Delivery Time')
corr_val = df[['distance_km','total_time_min']].corr().iloc[0,1]
slope, intercept, r_val, p_lr, se = stats.linregress(df['distance_km'], df['total_time_min'])
print(f'  Pearson r={corr_val:.3f}, slope={slope:.3f} min/km, R²={r_val**2:.3f}')

fig, axes = plt.subplots(1, 2, figsize=(16, 5))
fig.suptitle(f'Q5: Distance vs Delivery Time  (r={corr_val:.3f})', fontsize=15, fontweight='bold')
tc_map = {'Low':'#2A9D8F','Medium':'#E9C46A','High':'#F4A261','Very High':'#E63946'}
for traf, grp in df.groupby('traffic_level'):
    axes[0].scatter(grp['distance_km'], grp['total_time_min'],
                    color=tc_map.get(traf,'#888'), alpha=0.45, s=18, label=traf)
xr = np.linspace(df['distance_km'].min(), df['distance_km'].max(), 100)
axes[0].plot(xr, slope*xr+intercept, color='white', lw=2.5, label=f'y={slope:.2f}x+{intercept:.2f}')
axes[0].set_xlabel('Distance (km)'); axes[0].set_ylabel('Total Time (min)')
axes[0].set_title('Scatter (by Traffic)'); axes[0].legend(fontsize=9); axes[0].grid(True)
df['dist_bin'] = pd.cut(df['distance_km'], bins=[0,5,10,15,22], labels=['0-5km','5-10km','10-15km','15+km'])
d_bins = df.groupby('dist_bin', observed=True)['total_time_min'].mean()
axes[1].bar(d_bins.index, d_bins.values, color=['#2A9D8F','#E9C46A','#F4A261','#E63946'], edgecolor='#0F1117')
for i, v in enumerate(d_bins.values): axes[1].text(i, v+0.3, f'{v:.1f}', ha='center', fontsize=11)
axes[1].set_ylabel('Mean Total Time (min)'); axes[1].set_title('Avg Time by Distance Band')
axes[1].grid(True, axis='y')
plt.tight_layout(); savefig('fig08_distance_vs_time')

# Q6: Temporal congestion
print('\n  Q6: Temporal Congestion')
df['order_hour_num'] = df['order_time'].dt.hour
df['order_dow']      = df['order_time'].dt.day_name()
hourly = df.groupby('order_hour_num').agg(
    avg_time  = ('total_time_min','mean'),
    order_vol = ('order_id','count'),
    pct_late  = ('status', lambda x: (x.isin(['Late','Failed'])).mean()*100)
).reset_index()

fig, axes = plt.subplots(2, 2, figsize=(18, 10))
fig.suptitle('Q6: Temporal Congestion Analysis', fontsize=16, fontweight='bold')
axes[0,0].plot(hourly['order_hour_num'], hourly['avg_time'], color='#E63946', lw=2.5, marker='o', ms=5)
axes[0,0].fill_between(hourly['order_hour_num'], hourly['avg_time'], alpha=0.2, color='#E63946')
axes[0,0].set_xlabel('Hour'); axes[0,0].set_ylabel('Mean Total Time (min)')
axes[0,0].set_title('Avg Delivery Time by Hour')
axes[0,0].set_xticks(range(0,24,2)); axes[0,0].grid(True)

axes[0,1].bar(hourly['order_hour_num'], hourly['order_vol'], color='#457B9D', edgecolor='#0F1117')
axes[0,1].set_xlabel('Hour'); axes[0,1].set_ylabel('Orders')
axes[0,1].set_title('Order Volume by Hour')
axes[0,1].set_xticks(range(0,24,2)); axes[0,1].grid(True, axis='y')

day_order_list = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
heat = (df.groupby(['order_dow','order_hour_num'])['total_time_min']
          .mean().unstack(fill_value=0).reindex(day_order_list))
sns.heatmap(heat, ax=axes[1,0], cmap='YlOrRd', linewidths=0.3, cbar_kws={'label':'Avg Time (min)'})
axes[1,0].set_title('Heatmap: Delivery Time (Day × Hour)')

axes[1,1].bar(hourly['order_hour_num'], hourly['pct_late'], color='#E9C46A', edgecolor='#0F1117')
axes[1,1].set_xlabel('Hour'); axes[1,1].set_ylabel('% Late/Failed')
axes[1,1].set_title('Late Rate by Hour')
axes[1,1].set_xticks(range(0,24,2)); axes[1,1].grid(True, axis='y')
plt.tight_layout(); savefig('fig09_temporal_analysis')

peak_h = hourly.loc[hourly['avg_time'].idxmax(),'order_hour_num']
peak_v = hourly.loc[hourly['order_vol'].idxmax(),'order_hour_num']
print(f'  Peak congestion hour : {peak_h}:00')
print(f'  Peak order volume    : {peak_v}:00')

# Rider performance
rp = df.groupby('experience_level').agg(
    avg_time   = ('total_time_min','mean'),
    avg_rating = ('avg_rating','mean'),
    pct_late   = ('status', lambda x: (x.isin(['Late','Failed'])).mean()*100)
).reindex(['Junior','Mid','Senior'])

fig, axes = plt.subplots(1, 3, figsize=(18, 5))
fig.suptitle('Rider Experience vs Performance', fontsize=15, fontweight='bold')
exp_levels = ['Junior','Mid','Senior']; exp_colors = ['#E63946','#E9C46A','#2A9D8F']
for ax, col, ttl, ylab in zip(axes,
    ['avg_time','avg_rating','pct_late'],
    ['Avg Delivery Time','Avg Rating','Late/Failed Rate'],
    ['Mean Total Time (min)','Rating (out of 5)','Late/Failed %']):
    ax.bar(exp_levels, rp[col], color=exp_colors, edgecolor='#0F1117', alpha=0.85)
    ax.set_title(ttl); ax.set_ylabel(ylab); ax.grid(True, axis='y')
axes[1].set_ylim(0, 5)
plt.tight_layout(); savefig('fig10_rider_performance')
print('\n  Rider Performance:'); print(rp.round(2).to_string())

# Correlation heatmap
num_cols = ['total_time_min','delivery_duration_min','pickup_wait_min',
            'distance_km','order_value_inr','temperature_c',
            'rainfall_mm','visibility_km','experience_years','avg_rating']
corr = df[num_cols].corr()
fig, ax = plt.subplots(figsize=(12, 9))
mask = np.triu(np.ones_like(corr, dtype=bool))
sns.heatmap(corr, mask=mask, annot=True, fmt='.2f', ax=ax, cmap='coolwarm', center=0,
            annot_kws={'size':8}, linewidths=0.5, linecolor='#0F1117')
ax.set_title('Correlation Matrix – Key Variables')
plt.tight_layout(); savefig('fig11_correlation_heatmap')
print('\n  Correlations with total_time_min:')
print(corr['total_time_min'].drop('total_time_min').sort_values(key=abs, ascending=False).round(3).to_string())

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 9: FEATURE ENGINEERING
# ══════════════════════════════════════════════════════════════════════════════
print('\n[9] Feature Engineering...')

df['order_hour_num']    = df['order_time'].dt.hour
df['order_day_of_week'] = df['order_time'].dt.dayofweek
df['order_month']       = df['order_time'].dt.month
df['is_weekend']        = (df['order_day_of_week'] >= 5).astype(int)
df['is_peak_hour']      = df['order_hour_num'].apply(lambda h: 1 if (12<=h<=14) or (18<=h<=21) else 0)

def time_of_day(h):
    if 5<=h<12: return 'Morning'
    if 12<=h<17: return 'Afternoon'
    if 17<=h<21: return 'Evening'
    return 'Night'
df['time_of_day'] = df['order_hour_num'].apply(time_of_day)

traffic_num_map = {'Low':1,'Medium':2,'High':3,'Very High':4}
df['traffic_numeric']  = df['traffic_level'].map(traffic_num_map)
df['congestion_score'] = df['distance_km'] * df['traffic_numeric']
df['speed_kmpm']       = df['distance_km'] / df['delivery_duration_min'].replace(0, np.nan)
df['speed_kmpm'].fillna(df['speed_kmpm'].median(), inplace=True)

df['bad_weather']      = ((df['rainfall_mm'] > 10) | (df['visibility_km'] < 3)).astype(int)
df['weather_severity'] = ((df['rainfall_mm'].clip(0,30)/30)*50 + ((10-df['visibility_km'].clip(0,10))/10)*50).round(2)

exp_score_map = {'Junior':1,'Mid':2,'Senior':3}
df['rider_exp_score']       = df['experience_level'].map(exp_score_map)
df['rider_performance_idx'] = ((df['avg_rating']/5)*50 + (df['rider_exp_score']/3)*50).round(2)

rest_stats = df.groupby('restaurant').agg(
    rest_avg_pickup  = ('pickup_wait_min','mean'),
    rest_order_count = ('order_id','count'),
    rest_late_rate   = ('status', lambda x: (x.isin(['Late','Failed'])).mean())
).round(3)
df = df.merge(rest_stats, on='restaurant', how='left')

zone_stats = df.groupby('zone').agg(
    zone_avg_time  = ('total_time_min','mean'),
    zone_late_rate = ('status', lambda x: (x.isin(['Late','Failed'])).mean()),
    zone_order_vol = ('order_id','count')
).round(3)
df = df.merge(zone_stats, on='zone', how='left')

SLA = 45
df['is_delayed'] = (df['total_time_min'] > SLA).astype(int)

eng_features = ['delivery_duration_min','pickup_wait_min','total_time_min',
                'order_hour_num','order_day_of_week','order_month','is_weekend','is_peak_hour','time_of_day',
                'traffic_numeric','congestion_score','speed_kmpm',
                'bad_weather','weather_severity',
                'rider_exp_score','rider_performance_idx',
                'rest_avg_pickup','rest_order_count','rest_late_rate',
                'zone_avg_time','zone_late_rate','zone_order_vol','is_delayed']
print(f'  {len(eng_features)} features engineered')
print(f'  Delayed orders: {df["is_delayed"].sum()} ({df["is_delayed"].mean():.1%})')

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 10: ENCODING
# ══════════════════════════════════════════════════════════════════════════════
print('\n[10] Encoding categorical variables...')
tod_map = {'Morning':0,'Afternoon':1,'Evening':2,'Night':3}
df['time_of_day_enc'] = df['time_of_day'].map(tod_map)

le_status  = LabelEncoder()
le_wc      = LabelEncoder()
df['status_enc']            = le_status.fit_transform(df['status'])
df['weather_condition_enc'] = le_wc.fit_transform(df['weather_condition'])

ohe_cols = ['zone','restaurant','shift']
df_encoded = pd.get_dummies(df, columns=ohe_cols, prefix=ohe_cols, drop_first=True, dtype=int)
new_ohe = [c for c in df_encoded.columns if any(c.startswith(p+'_') for p in ohe_cols)]
print(f'  One-hot encoding: {len(new_ohe)} new columns')
print(f'  Encoded dataset shape: {df_encoded.shape}')

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 11: FEATURE SELECTION
# ══════════════════════════════════════════════════════════════════════════════
print('\n[11] Feature Selection...')
EXCLUDE = ['delivery_id','order_id','rider_id','route_id',
           'order_time','pickup_time','delivered_time','order_hour_key',
           'total_time_min','delivery_duration_min','pickup_wait_min',
           'status','status_enc','is_delayed',
           'time_of_day','weather_condition','dist_bin','rain_category','order_dow']

feature_cols = [c for c in df_encoded.select_dtypes(include=[np.number]).columns if c not in EXCLUDE]
X = df_encoded[feature_cols].fillna(df_encoded[feature_cols].median())
y = df_encoded['total_time_min']
print(f'  Feature matrix: {X.shape}')

# SelectKBest
skb = SelectKBest(f_regression, k='all')
skb.fit(X, y)
skb_scores = pd.Series(skb.scores_, index=feature_cols).sort_values(ascending=False)
print('\n  Top 15 by F-Regression:')
print(skb_scores.head(15).round(2).to_string())

# Random Forest
rf = RandomForestRegressor(n_estimators=150, max_depth=8, random_state=42, n_jobs=-1)
rf.fit(X, y)
rf_imp = pd.Series(rf.feature_importances_, index=feature_cols).sort_values(ascending=False)
print('\n  Top 15 by Random Forest:')
print(rf_imp.head(15).round(4).to_string())

# Feature importance plot
fig, axes = plt.subplots(1, 2, figsize=(18, 7))
fig.suptitle('Feature Importance – Predicting Total Delivery Time', fontsize=15, fontweight='bold')
c_fi = [PALETTE[i % len(PALETTE)] for i in range(15)]
t15r = rf_imp.head(15); t15s = skb_scores.head(15)
axes[0].barh(t15r.index[::-1], t15r.values[::-1], color=c_fi[::-1], edgecolor='#0F1117', alpha=0.85)
axes[0].set_xlabel('Importance'); axes[0].set_title('Random Forest Importance (Top 15)')
axes[0].grid(True, axis='x')
axes[1].barh(t15s.index[::-1], t15s.values[::-1], color=c_fi[::-1], edgecolor='#0F1117', alpha=0.85)
axes[1].set_xlabel('F-Score'); axes[1].set_title('SelectKBest F-Scores (Top 15)')
axes[1].grid(True, axis='x')
plt.tight_layout(); savefig('fig12_feature_importance')

# PCA
scaler = StandardScaler()
X_sc = scaler.fit_transform(X)
pca  = PCA(random_state=42); pca.fit(X_sc)
cumvar = np.cumsum(pca.explained_variance_ratio_)
n90 = np.argmax(cumvar >= 0.90) + 1
n95 = np.argmax(cumvar >= 0.95) + 1

fig, ax = plt.subplots(figsize=(10, 5))
ax.plot(range(1,len(cumvar)+1), cumvar*100, color='#457B9D', lw=2.5, marker='o', ms=4)
ax.axhline(90, color='#E9C46A', ls='--', lw=1.5, label=f'90% ({n90} components)')
ax.axhline(95, color='#E63946', ls='--', lw=1.5, label=f'95% ({n95} components)')
ax.fill_between(range(1,len(cumvar)+1), cumvar*100, alpha=0.15, color='#457B9D')
ax.set_xlabel('Components'); ax.set_ylabel('Cumulative Variance (%)')
ax.set_title('PCA Scree Plot'); ax.legend(); ax.grid(True)
plt.tight_layout(); savefig('fig13_pca_scree')
print(f'  PCA: {n90} components → 90% variance  |  {n95} components → 95% variance')

# Final selected features
top_rf  = set(rf_imp.head(12).index)
top_skb = set(skb_scores.head(12).index)
selected = sorted(top_rf.union(top_skb))
print(f'\n  Final selected features ({len(selected)}):')
for f in selected:
    rr = list(rf_imp.index).index(f)+1 if f in rf_imp.index else 'N/A'
    sr = list(skb_scores.index).index(f)+1 if f in skb_scores.index else 'N/A'
    print(f'    {f:45s}  RF={rr}  F={sr}')

# ══════════════════════════════════════════════════════════════════════════════
# SAVE OUTPUT FILES
# ══════════════════════════════════════════════════════════════════════════════
df.to_csv('master_cleaned.csv', index=False)
df_encoded.to_csv('master_encoded.csv', index=False)
print(f'\n  Saved master_cleaned.csv: {df.shape}')
print(f'  Saved master_encoded.csv: {df_encoded.shape}')

# ══════════════════════════════════════════════════════════════════════════════
# FINAL SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
print('''
╔══════════════════════════════════════════════════════════════════╗
║   QuickBite Delivery – Round 2 COMPLETE                         ║
╠══════════════════════════════════════════════════════════════════╣
║  7 data quality issues fixed                                     ║
║  5 datasets merged into 1 master dataset                         ║
║  6 use-case analytical questions answered with visuals           ║
║  23 new features engineered                                      ║
║  3 feature selection methods applied                             ║
║  13 visualizations saved as PNG                                  ║
║  master_cleaned.csv + master_encoded.csv saved                   ║
╚══════════════════════════════════════════════════════════════════╝
''')
