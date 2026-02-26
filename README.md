# 📈 Automated Stock Market Dashboard  
### Google Finance → Apps Script → Google Sheets → Power BI

An end-to-end automated Business Intelligence project that fetches daily stock market data from Google Finance, stores it in Google Sheets using Google Apps Script (incremental loading), and visualizes it in an interactive Power BI dashboard with dynamic time filtering (3M, 6M, 1Y, YTD, etc.).

---

## 🔗 Live Project Links

📊 **Power BI Public Dashboard:**  
https://app.fabric.microsoft.com/view?r=eyJrIjoiNWZiNjg5NTYtOTg2YS00Yzg5LWE5OTYtODAxMjYwY2FjNDJlIiwidCI6ImYzNmNjZTU5LTZlOWMtNDM3Ni1hZGFjLWE2MTVmMWIwMTBjNCJ9

📄 **Google Sheets Data Source:**  
https://docs.google.com/spreadsheets/d/1xvDkmo4i83m0krWpNj6d37BGiaqY8FsUCgemvPUj554/

---

# 🚀 Project Overview

This project demonstrates how to build a fully automated stock analytics pipeline using:

- Google Finance (Data Source)
- Google Apps Script (Automation + Incremental ETL)
- Google Sheets (Cloud Data Storage)
- Power BI Desktop (Data Modeling + Visualization)
- Power BI Service (Web Publishing)

The system automatically updates stock data daily and provides dynamic time-based filtering inside Power BI.

---

# 🏗️ Architecture
Google Finance
↓
Google Apps Script (Daily Trigger)
↓
Google Sheets (Cloud Storage)
↓
Power BI (Data Model + DAX)
↓
Published Web Dashboard


This lightweight architecture avoids paid APIs while maintaining automation, scalability, and clean data modeling.

---

# 📊 Google Sheets Structure

The spreadsheet contains three tabs:

## 1️⃣ Sheet1 — Ticker Master Table

| Column     | Description |
|------------|------------|
| Symbol     | Stock ticker (e.g., AAPL, MSFT) |
| Type       | Asset type |
| Logo_url   | Logo reference |

**Purpose:**  
Controls which stocks are tracked.  
Add a new ticker → system updates automatically.

---

## 2️⃣ HistoricalData — Fact Table

| Column     | Description |
|------------|------------|
| Symbol     | Stock ticker |
| Type       | Asset type |
| Date       | Trading date |
| Open       | Opening price |
| High       | Highest price |
| Low        | Lowest price |
| Close      | Closing price |
| Adj Close  | Adjusted close |
| Volume     | Trading volume |

This is the primary dataset consumed by Power BI.

---

## 3️⃣ Temp — Processing Sheet

Used temporarily by Apps Script to:

- Inject `GOOGLEFINANCE()` formula
- Extract computed values
- Clear after processing

---

# ⚙️ Automation with Google Apps Script

Main function:

```javascript
function updateHistoricalData()
```

What It Does

- Reads ticker list from Sheet1
- Checks the last available date per stock
- Fetches only missing data (Incremental Load)
- Dynamically builds GOOGLEFINANCE() formula
- Appends new rows in bulk
- Runs automatically via daily trigger

Key Strengths

- Prevents duplicate data
- Efficient incremental loading
- Fully automated
- Easily scalable for multiple stocks
- This acts as a lightweight ETL pipeline.

📐 Power BI Data Model

The Power BI report contains 4 tables:

## 1️⃣ Sheet1
Ticker metadata table.

## 2️⃣ HistoricalData
Main fact table (daily stock data).

## 3️⃣ DateTable (Custom Calendar Table)
```
DateTable = 
ADDCOLUMNS(
    CALENDAR (DATE(2025,1,1), TODAY()),
    "Year", YEAR([Date]),
    "Month", FORMAT([Date], "MMM"),
    "Month No", MONTH([Date]),
    "YearMonth", FORMAT([Date], "YYYY-MM")
)
```

Purpose:

- Enables time intelligence
- Clean X-axis behavior
- Proper star-schema modeling

Relationship:
```
DateTable[Date] → HistoricalData[Date]
```

## 4️⃣ PeriodSelector (Disconnected Table)
```
PeriodSelector = 
DATATABLE(
    "Period", STRING,
    {
        {"3M"},
        {"6M"},
        {"9M"},
        {"1Y"},
        {"5Y"},
        {"YTD"},
        {"All"}
    }
)
```

Used as a slicer to dynamically control the time period displayed in the chart.

# 🧠 Core DAX Logic (Dynamic Time Filtering)

```
Closing Price Filtered = 
VAR SelectedPeriod = SELECTEDVALUE(PeriodSelector[Period], "All")
VAR MaxDate =
    CALCULATE(
        MAX(HistoricalData[Date]),
        ALL(HistoricalData),
        HistoricalData[Symbol] = SELECTEDVALUE(HistoricalData[Symbol])
    )
VAR MinDate =
    SWITCH(
        SelectedPeriod,
        "3M", EDATE(MaxDate, -3),
        "6M", EDATE(MaxDate, -6),
        "9M", EDATE(MaxDate, -9),
        "1Y", EDATE(MaxDate, -12),
        "5Y", EDATE(MaxDate, -60),
        "YTD", DATE(YEAR(MaxDate),1,1),
        "All", DATE(2025,1,1)
    )
RETURN
IF(
    MAX(HistoricalData[Date]) >= MinDate &&
    MAX(HistoricalData[Date]) <= MaxDate,
    MAX(HistoricalData[Close]),
    BLANK()
)
```
### What This Enables

- Dynamic period selection (3M, 6M, 1Y, YTD, etc.)
- Filtering relative to latest available stock date
- Preserves daily granularity
- Clean and responsive stock trend visualization

# 📈 Dashboard Features

- Dynamic Time Period Selection (3M, 6M, 9M, 1Y, 5Y, YTD, All)
- Interactive Line Chart (Daily Closing Prices)
- Stock Symbol Slicer
- Proper Calendar Table
- Public Web Publishing
- Automated Daily Data Refresh

# 🛠️ Tech Stack

- Google Finance
- Google Apps Script
- Google Sheets
- Power BI Desktop
- Power BI Service
- DAX
- Star Schema Data Modeling

# 🎯 Skills Demonstrated

- Automated data pipelines
- Incremental ETL design
- Data modeling best practices
- Time intelligence in Power BI
- Disconnected slicer pattern
- Dynamic DAX filtering
- Cloud BI deployment

# 🔮 Future Enhancements

- Percentage return calculation
- Moving averages (MA20, MA50, MA200)
- Multi-stock comparison view
- Sector classification
- Financial ratios integration
- Scheduled Power BI Service refresh

# 📌 Conclusion

This project demonstrates how to build a scalable, automated, and professional-grade stock analytics dashboard using free and accessible tools.

It combines:

- Automation
- Data engineering concepts
- BI modeling
- Advanced DAX logic
- Cloud deployment

### If you found this project insightful, feel free to ⭐ the repository.
