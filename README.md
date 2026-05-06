# ⚡ Energy Consumption Data Analyst

A **production-ready Streamlit application** for importing, converting, analysing, and visualising electricity consumption data from Excel files.

---

## 🎯 Features

| Feature | Description |
|---|---|
| **15 min → 1 hour** | Sum 15-minute kWh readings into hourly totals |
| **1 hour → 15 min** | Expand hourly readings into 15-minute intervals (÷ 4) |
| **Leap-day removal** | Automatic detection and optional removal of Feb 29 data |
| **Finnish date parsing** | Handles Finnish weekday names (e.g. *Maanantai 01.01.2023 00:00*) |
| **Price column support** | Optional 3rd column with €/kWh for cost analysis |
| **Interactive charts** | Time-series, daily/weekly profiles, heatmaps, box-plots, cumulative curves |
| **Statistical summary** | Mean, median, std, percentiles, daily aggregates |
| **Excel export** | One-click download of processed data |

---

## 📁 Expected Excel File Structure

The app expects Excel files (`.xlsx` / `.xls`) with the following structure:

| Column A (DateTime) | Column B (kWh) | Column C (Price, optional) |
|---|---|---|
| Maanantai 01.01.2023 00:00 | 1.234 | 0.085 |
| Maanantai 01.01.2023 00:15 | 0.987 | 0.085 |
| ... | ... | ... |

- **Column A**: Finnish weekday name + date + time (e.g. `Maanantai 01.01.2023 00:15`)
- **Column B**: Energy consumption in kWh
- **Column C** *(optional)*: Price per kWh in EUR

---

## 🚀 Quick Start

### 1. Install dependencies

```bash
# Using uv (recommended)
uv sync

# Or using pip
pip install -r requirements.txt
```

### 2. Run the application

```bash
# Using uv
uv run streamlit run analyst.py

# Or directly
streamlit run analyst.py
```

### 3. Use the app

1. **Upload** your Excel file via the sidebar
2. **Select** the conversion operation (or keep original resolution)
3. **Toggle** leap-day removal if needed
4. **Explore** the auto-generated charts and statistics
5. **Download** the processed file as Excel

---

## 📊 Available Visualisations

- **Time-Series Plot** – Full consumption curve with filled area
- **Daily Load Profile** – Average kWh by hour of day
- **Weekly Load Profile** – Average kWh by day of week
- **Monthly Totals** – Bar chart of monthly consumption
- **Consumption Heatmap** – Hour × Day-of-week intensity map
- **Monthly Box-Plot** – Daily consumption distribution per month
- **Cumulative Curve** – Running total of energy consumed
- **Price & Cost Analysis** – Dual-axis chart (requires price column)

---

## 🛠️ Technology Stack

- **[Streamlit](https://streamlit.io/)** – Web application framework
- **[Pandas](https://pandas.pydata.org/)** – Data manipulation
- **[Plotly](https://plotly.com/python/)** – Interactive visualisations
- **[OpenPyXL](https://openpyxl.readthedocs.io/)** – Excel file I/O
- **[NumPy](https://numpy.org/)** – Numerical operations

---

## 📜 License

Internal tool – © 2026
