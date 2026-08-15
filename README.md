# EV Charging Network Usage Analysis | Microsoft Excel

## 📊 Project Overview

This project analyses **6,500 EV charging sessions across 8 cities** using Microsoft Excel to understand charging behaviour, revenue performance, station utilisation, vehicle usage and customer experience.

The project demonstrates an end-to-end Excel analytics workflow, including **data cleaning, formula-driven analysis, statistical analysis, PivotTables, visualisation and interactive dashboard development**.

---

## 🎯 Business Objective

The objective of this analysis was to transform raw EV charging-session data into meaningful business insights by answering questions such as:

- How much revenue is generated from charging sessions?
- Which cities generate the highest revenue?
- Which station type is used most frequently?
- Which vehicle types consume the most energy?
- How variable is charging-session revenue?
- What is the relationship between charging duration, energy consumption and revenue?
- How satisfied are customers with the charging experience?
- How can these insights be presented through an interactive Excel dashboard?

---

## 🗂️ Dataset

The dataset contains **6,500 EV charging sessions** across multiple cities.

### Key Columns

| Column | Description |
|---|---|
| Session_ID | Unique charging session identifier |
| City | City where charging occurred |
| Station_Type | Type of charging station |
| Vehicle_Type | EV type |
| Membership_Type | Member / Non-member |
| Charging_Start | Charging start date and time |
| Charging_End | Charging end date and time |
| Charging_Duration_Minutes | Total charging duration |
| Energy_Consumed_kWh | Electricity consumed |
| Price_per_kWh | Price charged per kWh |
| Revenue | Revenue generated |
| Payment_Mode | Payment method |
| Customer_Rating | Customer satisfaction rating |

---

# 🛠️ Tools & Excel Skills

### Microsoft Excel

- Data Cleaning & Validation
- Excel Tables
- Conditional Formatting
- PivotTables
- PivotCharts
- Dashboard Development
- Slicers
- KPI Cards

### Excel Functions Used

- `IF`
- `DATE`
- `TIME`
- `AVERAGE`
- `AVERAGEIF`
- `SUM`
- `COUNTA`
- `VLOOKUP`
- `INDEX-MATCH`
- `MODE`
- `STDEV.S`
- `VAR.S`
- `CORREL`
- ### Excel Functions Used

`IF` • `DATE` • `TIME` • `AVERAGE` • `AVERAGEIF` • `SUM` • `COUNTA` • `VLOOKUP` • `INDEX-MATCH` • `MODE` • `STDEV.S` • `VAR.S` • `CORREL`

---

---

# 🔄 Analysis Workflow

```text
Raw Dataset
     ↓
Data Cleaning & Validation
     ↓
Calculated Columns
     ↓
Lookup Analysis
     ↓
Statistical Analysis
     ↓
PivotTable Analysis
     ↓
Visualisations
     ↓
Interactive Dashboard
     ↓
Business Insights
1. 🧹 Data Cleaning & Validation

Before performing the analysis, the dataset was checked and prepared to ensure data quality and consistency.

Cleaning & Validation Steps
Checked for missing values
Checked for duplicate records
Verified numeric columns for valid numerical values
Verified date and time formats
Validated the Revenue calculation
Reviewed the dataset for consistency before analysis
Revenue Validation

Revenue was validated using:

Revenue = Energy Consumed × Price per kWh

Excel calculation:

=Energy_Consumed_kWh*Price_per_kWh

This validation helped ensure that the revenue values used in the analysis were consistent with the underlying charging data.
2. 🧮 Data Transformation & Calculated Columns

Additional calculated columns were created using Excel formulas to support further analysis.

Charging Time Category

Charging sessions were classified based on charging duration.

Charging Duration	Category
< 30 minutes	Short Charging
30–60 minutes	Medium Charging
> 60 minutes	Long Charging
Energy Usage Level

Energy consumption was classified into three levels.

Energy Consumed	Category
< 10 kWh	Low
10–30 kWh	Medium
> 30 kWh	High
Customer Satisfaction Level

Customer ratings were classified into satisfaction categories.

Customer Rating	Category
1–2	Low Satisfaction
3–4	Medium Satisfaction
5	High Satisfaction
Date & Time Extraction

Separate Date and Time columns were extracted from the Charging_Start timestamp.

Date
=DATE(YEAR(F2),MONTH(F2),DAY(F2))
Time
=TIME(HOUR(F2),MINUTE(F2),SECOND(F2))

These fields were used for time-based analysis and energy consumption trends.

3. 🔎 Lookup Analysis

A separate city-level lookup table was created to calculate and retrieve:

Average Energy Consumption
Average Charging Duration
Functions Used
=AVERAGEIF()
=VLOOKUP()

INDEX-MATCH and MODE were also used to identify the most frequently used station type.

4. 📈 Statistical Analysis

Descriptive statistical analysis was performed using Excel statistical functions.

Average Energy Consumption

Average energy consumed per charging session:

30.91 kWh
Revenue Variation

Coefficient of Variation:

60.45%
Interpretation

The coefficient of variation indicates substantial variability in revenue generated across individual charging sessions.

Variance Comparison
Variable	Variance
Energy Consumed	281.61
Price per kWh	6.76
Finding

Energy consumption showed considerably greater absolute variability than Price per kWh.

5. 🔗 Correlation Analysis

Correlation analysis was performed using Excel's CORREL() function.

Charging Duration vs Energy Consumed
Correlation = -0.010

This indicates essentially no linear relationship between charging duration and energy consumption in the analysed dataset.

Energy Consumed vs Revenue
Correlation = 0.900

This indicates a very strong positive relationship between energy consumed and revenue.

The relationship is consistent with:

Revenue = Energy Consumed × Price per kWh
6. 📊 PivotTable Analysis

PivotTables were created to analyse different aspects of EV charging behaviour.

Station Utilisation
Count of charging sessions by Station Type
Average Charging Duration by Station Type
Vehicle Usage
Energy Consumption by Vehicle Type
Revenue generated by Vehicle Type
Customer Experience
Average Customer Rating by Station Type
Rating Distribution by Membership Type
7. 📊 Visualisation

The project includes multiple Excel visualisations.

Charts Created
Revenue by City — Bar Chart
Average Energy Consumption by Vehicle Type — Column Chart
Charging Duration vs Energy Consumption — Scatter Plot
Station Type Distribution — Pie Chart
Energy Consumption Trends — Line/Trend Chart
8. 🎨 Conditional Formatting

Conditional formatting was applied to highlight important values.

Rules Applied
Highlight Revenue in the Top 10%
Highlight Charging Duration greater than 90 minutes
Apply a colour scale to Customer Rating
9. 📊 Interactive Excel Dashboard

An interactive Excel dashboard was created to provide a consolidated view of the analysis.

KPI Metrics
KPI	Result
Total Revenue	₹2,260,491.67
Total Charging Sessions	6,500
Average Energy Consumed	30.91 kWh
Average Customer Rating	2.99 / 5
Dashboard Visualisations
Revenue by City
Vehicle Type Usage
Energy Consumption Trends
Station Type Distribution
Charging Duration vs Energy Consumption
Interactivity

A City slicer was added to allow users to filter the dashboard by city.

10. 💡 Key Business Findings
Revenue Performance

Total revenue generated:

₹2,260,491.67

Highest revenue-generating city:

Kolkata — ₹308,122.98

Second-highest:

Hyderabad — ₹294,720.95
Station Utilisation

The most frequently used station type was:

Super Fast Charger

Sessions:

2,235

Share of total sessions:

34.38%
Vehicle Type Performance

3W EVs generated the highest vehicle-type revenue:

₹583,503.21

They also recorded the highest average energy consumption:

31.71 kWh/session
Customer Experience

Average customer rating:

2.99 / 5

Average rating by station type:

Station Type	Average Rating
Fast Charger	3.01
Standard Charger	3.00
Super Fast Charger	2.97
Revenue Variability

Revenue showed substantial variability with a:

60.45% Coefficient of Variation
Energy & Revenue Relationship

Energy consumption and revenue showed a strong positive relationship:

Correlation = 0.900
Charging Duration & Energy Relationship

Charging duration and energy consumption showed almost no linear relationship:

Correlation = -0.010
11. 📌 Business Recommendations

Based on the analysis:

Monitor Super Fast Charger capacity because it has the highest session utilisation.
Investigate customer experience because the overall average rating is around 3/5.
Monitor energy consumption as a key revenue driver when evaluating charging performance.
Investigate high- and low-revenue sessions because revenue variability is substantial.
Investigate charging behaviour further because charging duration and energy consumption showed almost no linear relationship in this dataset.



🧠 Skills Demonstrated

This project demonstrates practical experience in:

Data Analysis
Data Cleaning
Data Validation
Data Transformation
Descriptive Statistics
Correlation Analysis
Business Analysis
Business Insight Generation
Excel
Advanced Excel Formulas
PivotTables
PivotCharts
Conditional Formatting
KPI Development
Data Visualisation
Dashboard Development
Slicers
📌 Resume Project Description
EV Charging Network Usage Analysis | Microsoft Excel

Analysed 6,500 EV charging sessions across 8 cities using advanced Excel formulas, PivotTables, statistical analysis and dashboard visualisation; identified ₹2.26M total revenue, 60.45% revenue variability, and a 0.90 correlation between energy consumption and revenue.

Technical Skills

Microsoft Excel | Data Cleaning | Data Transformation | PivotTables | Statistical Analysis | VLOOKUP | INDEX-MATCH | AVERAGEIF | Conditional Formatting | Data Visualisation | Dashboard Development

🎯 Project Outcome

This project demonstrates the ability to take a raw operational dataset and transform it into a structured analytical solution using Microsoft Excel.

The complete workflow covers:

Data Cleaning → Transformation → Lookup Analysis → Statistical Analysis → PivotTables → Visualisation → Dashboard → Business Insights

The project demonstrates practical Excel and analytical skills relevant to an entry-level Data Analyst role.

👩‍💻 Author

Shriya Verma

Aspiring Data Analyst

Skills: Excel | SQL | Power BI | Data Analysis



