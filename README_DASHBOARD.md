# 📊 Sales Forecast Dashboard

Professional Excel dashboard generator for sales forecasting and variance analysis.

## 📁 Project Structure

```
rutik/
├── 3-3+forecast_table_raw.csv          # Source forecast data
├── forecast_dashboard.py               # Dashboard generator script
├── requirements.txt                    # Python dependencies
└── Sales_Forecast_Dashboard.xlsx       # Generated output (created after running script)
```

## 🚀 Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the Script
```bash
python forecast_dashboard.py
```

### 3. Output
Creates `Sales_Forecast_Dashboard.xlsx` with professional formatting and charts.

---

## 📊 Dashboard Components

### **Sheet 1: Raw Data**
Original CSV import with all metrics and drivers

### **Sheet 2: Summary** ⭐ Main Dashboard
| Column | Description |
|--------|-------------|
| **Metric** | Financial metrics (Revenue, COGS, Opex, EBITDA, Cash) |
| **Q1_Actual** | Actual Q1 performance |
| **Q2_Plan** | Q2 target plan |
| **Q2_Forecast** | Q2 forecasted values |
| **Plan_vs_Forecast** | Variance in dollars |
| **Variance_%** | Variance as percentage |
| **P80_Low/High** | Confidence interval bounds |

**Includes:**
- ✅ Currency formatting ($)
- ✅ Percentage calculations
- ✅ Professional styling (blue headers, borders)
- ✅ Bar chart: Revenue comparison
- ✅ Line chart: EBITDA trend

### **Sheet 3: Drivers**
Breakdown of key drivers affecting each metric:

| Metric | Driver_1 | Driver_2 | Driver_3 | Impact_Summary |
|--------|----------|----------|----------|---|
| **Revenue** | Price: +$2,000 | Volume: -$2,500 | FX: -$1,500 | Net: -$2,000 |
| **COGS** | Materials: +$1,500 | Efficiency: -$500 | FX: +$1,000 | Net: +$2,000 |
| **Opex** | Hiring: +$700 | Marketing: -$200 | — | Net: +$500 |

---

## 📈 Key Metrics at a Glance

| Metric | Q1 Actual | Q2 Plan | Q2 Forecast | Variance | Status |
|--------|-----------|---------|-------------|----------|--------|
| **Revenue** | $120K | $130K | $128K | -$2K (-1.5%) | ⚠️ Below Plan |
| **COGS** | $72K | $76K | $77K | +$1K (+1.3%) | ⚠️ Higher |
| **Opex** | $28K | $29.5K | $29K | -$0.5K (-1.7%) | ✅ Under Control |
| **EBITDA** | $20K | $24.5K | $22K | -$2.5K (-10.2%) | ⚠️ Below Target |
| **Cash_End** | $12K | $12.5K | $11.4K | -$1.1K (-8.8%) | ⚠️ Tight |

---

## 🔧 Features

✅ **Automatic Formatting**
- Professional color scheme
- Currency and percentage number formats
- Borders and alignment

✅ **Variance Analysis**
- Plan vs Forecast comparison
- P80 confidence intervals
- Percentage variance calculations

✅ **Driver Tracking**
- Price, Volume, FX impacts
- Materials, Efficiency, Opex breakdowns
- Impact summaries

✅ **Interactive Charts**
- Bar chart for revenue comparison
- Line chart for trend analysis
- Professional styling

---

## 📝 How to Use

1. **Update forecast data** in `3-3+forecast_table_raw.csv`
2. **Run the script**: `python forecast_dashboard.py`
3. **Open** `Sales_Forecast_Dashboard.xlsx`
4. **Analyze** trends and variances
5. **Share** with stakeholders

---

## 🛠️ Customization

Edit `forecast_dashboard.py` to:
- Add more metrics
- Include additional quarters
- Customize chart styles
- Add product category breakdowns
- Integrate with live data sources

---

## 📦 Requirements

- Python 3.7+
- pandas 2.0.3+
- openpyxl 3.1.2+

---

## 📧 Support

For issues or questions, check the generated Excel file or review the script comments.

---

**Version:** 1.0  
**Last Updated:** 2026-07-07
