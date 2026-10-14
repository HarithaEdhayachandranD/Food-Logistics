# 🛵 QuickBite Delivery Services — Food Delivery Insights & Analytics

> **Data Sprint Project**: End-to-end Data Science & Business Analytics Pipeline for Food Delivery Operations  
> **Target**: Understand delivery delays, identify operational bottlenecks, and enhance SLA fulfillment performance.

---

## 📁 Repository & Folder Organization

The project files have been neatly organized by project round and domain into the following directory structure:

```
Hari_DS_Spring/
├── README.md                                          <-- Complete project documentation & findings
├── Round_1_Data_Collection/
│   └── Raw_Datasets/
│       ├── orders.csv                                 <-- Order timestamp, restaurant, customer zone, order value
│       ├── riders.csv                                 <-- Rider ID, experience, shift, rating
│       ├── routes.csv                                 <-- Route distance (km), traffic level, zone
│       ├── weather.csv                                <-- Rain condition, temperature (°C), visibility (km)
│       └── delivery.csv                               <-- Pickup time, delivered time, delivery status
│
├── Round_2_EDA_Feature_Engineering/
│   ├── Cleaned_Data/
│   │   ├── master_cleaned.csv                         <-- Cleaned, imputed, merged 3,000-order dataset
│   │   └── master_encoded.csv                         <-- Machine learning-ready feature matrix
│   ├── Notebook/
│   │   └── QuickBite_Round2_EDA.ipynb                 <-- Comprehensive Jupyter Notebook for Round 2
│   ├── Scripts/
│   │   └── run_eda.py                                 <-- Standalone pipeline script for cleaning & stats
│   └── Visualizations/                                <-- 13 high-resolution diagnostic charts
│       ├── fig01_missing_heatmap.png
│       ├── fig02_delivery_time_dist.png
│       ├── fig03_status_distribution.png
│       ├── fig04_zone_delivery_time.png
│       ├── fig05_traffic_impact.png
│       ├── fig06_rain_impact.png
│       ├── fig07_restaurant_pickup.png
│       ├── fig08_distance_vs_time.png
│       ├── fig09_temporal_analysis.png
│       ├── fig10_rider_performance.png
│       ├── fig11_correlation_heatmap.png
│       ├── fig12_feature_importance.png
│       └── fig13_pca_scree.png
│
├── Round_3_Excel_Analysis/
│   ├── QuickBite_Round3_Excel_Analysis.xlsx           <-- Master 9-sheet Excel workbook with formulas & pivots
│   └── build_excel.py                                 <-- OpenPyXL automated generation engine
│
└── Scripts_Utilities/
    └── generate_data.py                               <-- Realistic dataset synthesizer with business noise
```

---

## 🔍 Round Summary & Evidence

### 📌 Round 1: Data Collection & Data Quality Ingestion
- **Context**: QuickBite coordinates multi-restaurant deliveries across urban and suburban customer zones.
- **Data Quality Ingestion & Fixes**:
  - Addressed missing timestamps and weather telemetry.
  - De-duplicated overlapping order records.
  - Standardized inconsistent zone naming conventions (e.g., casing, trailing whitespace).
  - Cleaned negative durations and chronological inversions.

---

### 📊 Round 2: Exploratory Data Analysis & Feature Engineering (30 Marks)
- **Cleaning & Consolidation**:
  - Merged 5 relational streams (`orders`, `riders`, `routes`, `weather`, `delivery`) on common keys (`order_id`, `rider_id`, `route_id`, `weather_id`).
  - Derived core duration metrics: `pickup_delay_mins`, `transit_time_mins`, and `total_delivery_time_mins`.
- **Engineered Features**:
  - `speed_kmh`: Calculated transit velocity.
  - `is_delayed`: SLA indicator flag (delivery time > 45 minutes).
  - `is_peak_hour`: Indicator for peak lunch (12:00–14:00) and dinner (19:00–22:00) windows.
  - `weather_severity_idx`: Composite index factoring rainfall and low visibility.
  - `traffic_encoded`: Ordinal encoding for Low, Moderate, High, Jam levels.
- **Visual Evidence**:
  - 13 figures saved under `Round_2_EDA_Feature_Engineering/Visualizations/`.

---

### 📈 Round 3: Data Analysis Using Microsoft Excel (30 Marks)
- **Workbook**: `Round_3_Excel_Analysis/QuickBite_Round3_Excel_Analysis.xlsx` (220 KB)
- **9 Structured Worksheets**:
  1. `🔴 Summary Dashboard`: Executive KPI scorecard, top analytical insights, and strategic recommendations.
  2. `🔵 Raw Data`: Full 3,000-row structured table with alternating row styles, status-based conditional formatting, and freeze panes.
  3. `📊 Statistical Summary`: Mean, Median, Mode, Std Dev, Variance, Min, Max, Skewness, Kurtosis, and correlation matrix (`CORREL`).
  4. `📍 Zone Analysis`: Aggregated performance by zone with Bar Chart.
  5. `🚦 Traffic Analysis`: Traffic level impact breakdown with Column Chart.
  6. `🌧️ Weather Impact`: Rain condition impact comparison with Bar Chart.
  7. `🍳 Restaurant Pickup`: Kitchen delay breakdown across food partners with Horizontal Bar Chart.
  8. `📏 Distance & Peaks`: Distance tier and peak vs off-peak comparison with dual charts.
  9. `🔄 Pivot Tables & Charts`: 3 Cross-tabulation pivot tables and a Stacked Bar Pivot Chart.

---

## 💡 Top Strategic Business Insights

1. **Traffic Congestion is Bottleneck #1**: "Jam" conditions cause average delivery durations to surge by **+78%** (from ~28 mins to ~50 mins).
2. **Rain Exacerbates Rider Availability & Speed**: Heavy rain increases transit times by an average of **+14.2 minutes** and lifts delay rates to over 42%.
3. **Restaurant Prep Variances**: Select high-volume restaurants have prep times exceeding 24 minutes, consuming over 50% of the customer SLA window.
4. **Zone-Specific Disparities**: Outer suburban zones have disproportionate transit delays due to sparse rider density and higher average route distances.
