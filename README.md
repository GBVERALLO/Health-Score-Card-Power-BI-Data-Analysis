# Health-Score-Card-Power-BI-Data-Analysis
An interactive Power BI dashboard that tracks the daily health indicators of 10 patients over 10 days and summarises them into a single, easy-to-read Health Score Card.

Project Overview
Item	Detail
Objective	Monitor patient wellness through one health score plus key vitals and lifestyle metrics
Granularity	One row per patient per day
Patients	10 (5 female, 5 male; ages 24–52)
Cities	Lagos, Abuja, Port Harcourt, Ibadan, Kaduna, Enugu, Benin City, Onitsha, Owerri, Accra
Period	1 Sep 2026 – 10 Sep 2026 (10 days)
Records	100 rows, 17 columns, no missing values
🗂️ Data Model

The Excel workbook has three sheets, each loaded as a table through Power Query.

Table	Purpose
HealthData	Main fact table (100 rows)
DAX Measures	Reference sheet that documents the measures
Dashboard Guide	Reference sheet that maps each dashboard component to a Power BI element

An auto-generated date table (LocalDateTable) supports the Year / Quarter / Month / Day hierarchy on the trend chart.

HealthData columns
Column	Type	Description
Patient ID	Text	Unique patient code (P001–P010)
Patient Name	Text	Patient's full name
Gender	Text	Male / Female
Age	Whole number	Age in years
City	Text	City of residence
Date	Date	Measurement date
Health Score	Whole number	Overall daily score (0–100)
Heart Rate (bpm)	Whole number	Resting heart rate
Systolic BP / Diastolic BP	Whole number	Blood pressure (mmHg)
BMI	Decimal	Body mass index
Sleep Hours	Decimal	Hours slept
Steps	Whole number	Daily step count
Calories Burned	Whole number	Daily calories burned
Stress Level	Whole number	Self-rated stress, 1 (low) – 10 (high)
Health Status	Text	Category label: Excellent or Good
Profile Image URL	Text	Patient photo, displayed in the report
🔧 Data Preparation (Power Query)
Connected to the Excel workbook with Excel.Workbook.
Selected the HealthData, DAX Measures and Dashboard Guide sheets.
Promoted the first row to headers.
Set data types: text for IDs, names and URLs; whole number for Age, Health Score, BP, Steps, Calories and Stress Level; decimal for BMI and Sleep Hours; date for Date.
Checked data quality: no nulls and 10 unique patients, each with 10 daily records.
🧮 DAX Measures
DAX
Average Health Score = AVERAGE(HealthData[Health Score])
Average Steps        = AVERAGE(HealthData[Steps])
Average Sleep hours  = AVERAGE(HealthData[Sleep Hours])
Average heart rate   = AVERAGE(HealthData[Heart Rate (bpm)])

Dynamic Patient Title =
"Health Overview — " & SELECTEDVALUE(HealthData[Patient Name], "All Patients")

Greetings =
VAR CurrentHour = HOUR(NOW())
VAR Greeting =
    SWITCH(
        TRUE(),
        CurrentHour < 12, "Good Morning",
        CurrentHour < 17, "Good Afternoon",
        "Good Evening"
    )
RETURN
    Greeting & " " & SELECTEDVALUE(HealthData[Patient Name]) & " 👋"

Health Insight =
VAR Score = [Average Health Score]
RETURN
SWITCH(
    TRUE(),
    Score >= 85, "Overall health indicators are strong.",
    Score >= 75, "Overall health indicators are good.",
    Score >= 65, "Some health indicators should be monitored.",
    "Several indicators require attention."
)
📊 Dashboard Components
Visual	Type	What it shows
Greeting	Card	Time-based welcome message with the selected patient's name
Health Score	KPI card	Average health score
Average Steps	KPI card	Activity level
Average Sleep Hours	KPI card	Sleep quality
Average Heart Rate	KPI card	Heart-rate overview
Health Score Trend	Line chart	Average score by Year → Quarter → Month → Day
Profile Image	Image visual	Photo of the selected patient
Patient Name	Slicer	Filters the whole report to one patient

The report is a single page at 1920 × 1080 with a green/white theme.

🔍 Key Findings
Overall (all 10 patients, 10 days)
Metric	Average	Min – Max
Health Score	85.6	76 – 96
Heart Rate	70.3 bpm	53 – 86
Blood Pressure	120.7 / 77.0 mmHg	102–140 / 64–92
BMI	24.0	18.1 – 32.8
Sleep	7.0 hrs	5.3 – 9.5
Steps	7,040 / day	3,952 – 10,486
Calories Burned	1,450 / day	1,118 – 1,764
Stress Level	4.9 / 10	1 – 10
Strong overall health. The average score of 85.6 falls in the "strong" band (≥ 85) of the Health Insight measure. 89% of records are Excellent (score ≥ 80) and 11% are Good (76–79). No record is below 76.
Top performers. Emeka Chukwu (88.9), Esther Danjuma (86.6) and Adaeze Nwosu (86.5) have the highest average scores. Daniel Adeyemi (83.4) and Grace Umeh (83.6) have the lowest, though still in the healthy range.
Steps are the strongest driver of the score. Steps have the highest correlation with Health Score (r ≈ 0.58). Heart rate (0.24), calories burned (0.17) and BMI (0.15) are weak, and sleep, stress and blood pressure are close to zero.
Daily trend. The daily average score ranged from 83.3 (2 Sep) to 88.4 (5 Sep). It rose early in the period, then eased back to 83.9 by 10 Sep.
Gender. Male and female averages are almost the same (85.8 vs 85.3).
Watch points.
About 49% of records have sleep under 7 hours.
About 45% of records have blood pressure at or above 130 systolic or 80 diastolic.
About 35% of records have a BMI of 25 or more.
Only about 4% of records fall below 5,000 steps.

Findings are computed from the 100 records in the dataset. The dataset appears to be sample or synthetic data (profile photos come from randomuser.me), so the results are for demonstration and learning, not clinical use.

▶️ How to Use
Install Power BI Desktop.
Open IFEXA-DA-0486-HEALTH_SCORE_CARD.pbix.
If Power BI cannot find the data, go to Transform data → Data source settings and point it at your local copy of Health_Scorecard_PowerBI_Dataset.xlsx.
Use the Patient Name slicer to switch between patients. The greeting, photo, KPI cards and trend chart update automatically.
🧰 Skills Demonstrated
Importing and cleaning Excel data with Power Query
Data typing and data quality checks
Writing DAX measures (aggregations, VAR/RETURN, SWITCH, SELECTEDVALUE)
Dynamic, time-aware text and conditional insight messages
KPI cards, line chart with date hierarchy, slicers and image visuals
Turning raw data into a clear, interactive report
📁 Repository Structure
├── IFEXA-DA-0486-HEALTH_SCORE_CARD.pbix
├── Health_Scorecard_PowerBI_Dataset.xlsx
├── screenshots/
│   └── dashboard.png
└── README.md

💡 Add a screenshot of the dashboard to screenshots/ and show it at the top with ![Dashboard](screenshots/dashboard.png).

👤 Author

Your Name · Gerald B. Verallo
