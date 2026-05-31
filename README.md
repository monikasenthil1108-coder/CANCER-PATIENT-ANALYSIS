🏥**Cancer Patient Data Analysis & Severity Prediction (2022–2024)**
> An end-to-end data analytics mini project using **Microsoft Excel** and **Power BI** to clean, analyse, and visualize global cancer patient data.
>
**> 📌 Project Title**
> Cancer Patient Data Analysis and Severity Prediction using Excel & Power BI

📋 **Problem Statemen**t

Cancer is a major global health issue affecting people across different ages, genders, and countries. This project aims to clean and analyse a cancer patient dataset containing 549 records collected from 10 countries over 3 years (2022–2024)with 15 attributes including risk factors, cancer type, stage, treatment cost, and survival years.

The raw dataset contains real-world data quality issues like missing values, duplicates, typos, and inconsistencies which are resolved through data cleaning. After cleaning, the data is analysed to understand how lifestyle risk factors affect cancer severity, how treatment costs vary across cancer types and countries, and how cancer stage impacts patient survival.

🎯 **Objectives**

- ✅ Clean and preprocess raw cancer patient data using Excel
- ✅ Handle missing values, duplicates, typos, and formatting issues
- ✅ Perform Pivot Table analysis for initial insights
- ✅ Build an interactive Power BI dashboard with DAX measures
- ✅ Derive meaningful insights on cancer severity, treatment cost, and survival patterns

📁 **Project Structure**


📦 **Cancer-Patient-Analysis**
 
 ┣ 📄 MINI_PROJECT_CLEAN → Cleaned dataset (549 rows)
 
 ┣ 📄 MINI_PROJECT_UNCLEANED → Uncleaned dataset (with issues)

 ┣ 📄 FINAL_MINI_PROJECT → Final cleaned Excel file

 ┣ 📄 MINI PROJECT REPORT → Full project report

 ┗ 📄 README.md → Project documentation

📊 **Dataset Overview**

 Attribute Details 

Source → Kaggle — Global Cancer Patient Dataset

Total Records → 549 rows (shortlisted from 14,810) 

Total Columns → 15 attributes 

Time Period → 2022, 2023, 2024 

Countries → USA, India, UK, Germany, Australia, Brazil, Canada, China, Pakistan, Russia 

Cancer Types → Lung, Skin, Breast, Liver, Prostate, Colon, Cervical, Leukemia 

Cancer Stages → Stage 0, Stage I, Stage II, Stage III, Stage IV 

📋 **Columns Description**

 Column | Data Type | Description 

 Patient_ID | String | Unique patient identifier 
 
 Age | Integer | Patient age in years 
 
 Gender | Categorical | Male / Female / Other 
 
 Country_Region | Categorical | Country of the patient 
 
 Year | Integer | Year of data collection 
 
 Genetic_Risk | Float | Genetic risk factor score (0–10) 
 
 Air_Pollution | Float | Air pollution exposure level (0–10) 
 
 Alcohol_Use | Float | Alcohol usage level (0–10) 
 
 Smoking | Float | Smoking level score (0–10) 
 
 Obesity_Level | Float | Obesity level score (0–10) 
 
 Cancer_Type | Categorical | Type of cancer diagnosed 
 
 Cancer_Stage | Categorical | Stage of cancer (Stage 0–IV) 
 
 Treatment_Cost_USD | Float | Treatment cost in USD 
 
 Survival_Years | Float | Years survived after treatment 
 
 Target_Severity_Score | Float | Overall cancer severity score 


**Tools & Technologies**

Tool-Purpose 

Microsoft Excel → Data cleaning, formula-based imputation, Pivot Tables 
 
Microsoft Power BI → DAX measures, visualizations, interactive dashboard 

GitHub → Version control and project documentation 

🧹 **Data Cleaning Steps (Excel)**

The uncleaned dataset contained 11 types of data quality issues. The following steps were performed to clean the data:

Issue-Method Used 

 1. Duplicate Records = Data → Remove Duplicates → Patient_ID column → 39 duplicates removed 

 2. Blank / Whitespace Cells | Find & Replace + Ctrl+G → Special → Blanks 
 
 3. Case Inconsistencies | `=PROPER()` formula on Cancer_Type, Country_Region, Gender, Cancer_Stage 
 
 4. Typos & Wrong Spellings | Find & Replace with Match Entire Cell Contents 
 
 5. Inconsistent Gender Entries | Find & Replace (MALE→Male, fem→Female, M→Male) 
 
 6. M.issing Cancer Type | Filled with 'Unknown' using Ctrl+G → Blanks
 
 7. Missing Risk Factors | Year-wise AVERAGEIF formula 
 
 8. Missing Severity Score | Average of 5 risk factor columns 
 
 9. Missing/Wrong Cancer Stage | Nested IF formula based on Severity Score 
 
 10.  Missing Survival Years | Stage-wise AVERAGEIF formula 
 
 11. Missing Treatment Cost | AVERAGEIFS by Cancer Type + Cancer Stage 
 
 12. Mixed Formatting ($ sign) | Changed Currency format to Number format 

**🔢 Key Formulas Used**

Excel

 1.Year-wise average for Risk Factors
=IF(H2<>"",H2,ROUND(AVERAGEIF($G$2:$G$532,G2,$H$2:$H$532),1))

 2.Severity Score from Risk Factors
=IF(R2<>"",R2,ROUND((F2+G2+H2+I2+J2)/5,2))

 3.Cancer Stage from Severity Score
=IF(W2<=2,"Stage 0",IF(W2<=4,"Stage I",IF(W2<=6,"Stage II",IF(W2<=8,"Stage III","Stage IV"))))

 4.Survival Years by Stage
=IF(V2<>"",V2,ROUND(AVERAGEIF($T$2:$T$532,T2,$V$2:$V$532),1))

 5.Treatment Cost by Cancer Type + Stage
=IF(U2<>"",U2,ROUND(AVERAGEIFS($U$2:$U$532,$S$2:$S$532,S2,$X$2:$X$532,X2),2))

## 📈 Pivot Tables (Excel)

| # | Pivot Table | Rows | Values |
|---|---|---|---|
| 1 | Cancer Type wise Patient Count | Cancer_Type | Count of Patient_ID |
| 2 | Year wise Avg Treatment Cost | Year | Average of Treatment_Cost_USD |
| 3 | Stage wise Avg Survival Years | Cancer_Stage | Average of Survival_Years |
| 4 | Country wise Patient Count | Country_Region | Count of Patient_ID |
| 5 | Gender wise Cancer Type Distribution | Cancer_Type / Gender | Count of Patient_ID |
| 6 | Cancer Type wise Avg Treatment Cost | Cancer_Type | Average of Treatment_Cost_USD |


## 💡 DAX Measures (Power BI)

```dax
-- Total Patients
Total Patients = COUNTROWS(Sheet1)

-- Average Treatment Cost
Avg Treatment Cost = AVERAGE(Sheet1[Treatment_Cost_USD])

-- Average Survival Years
Avg Survival Years = AVERAGE(Sheet1[Survival_Years])

-- Average Severity Score
Avg Severity Score = AVERAGE(Sheet1[TARGET_SEVERITY_SCORE])

-- High Risk Patients (Severity > 7)
High Risk Patients = COUNTROWS(FILTER(Sheet1, Sheet1[TARGET_SEVERITY_SCORE] > 7))
```

---

## 📊 Power BI Dashboard

### KPI Cards
| Metric | Value |
|---|---|
| Total Patients | 531 |
| Average Treatment Cost | $51.82K |
| Average Survival Years | 5.47 |
| Average Severity Score | 5.77 |
| High Risk Patients | 44 |

### Visualizations

| # | Chart Title | Chart Type |
|---|---|---|
| 1 | Cancer Type wise Patient Count | Bar Chart |
| 2 | Gender Distribution | Pie Chart |
| 3 | Cancer Stage wise Avg Survival Years | Column Chart |
| 4 | Year wise Avg Treatment Cost Trend | Line Chart |
| 5 | Country wise Patient Distribution | Treemap |
| 6 | Year wise Patient Count | Area Chart |
| 7 | Cancer Stage wise Patient Count | Donut Chart |
| 8 | Cancer Type wise Avg Severity Score | Funnel Chart |

### Slicers
- 📅 Year (2022 / 2023 / 2024)
- 🎗️ Cancer Type
- 📊 Cancer Stage

---

## 🔍 Key Insights
### 📌 Descriptive
- 531 patients across 10 countries and 3 years
- Colon cancer had the highest patient count (73); Cervical had the lowest (64)
- Gender distribution: Female (35%), Male (35%), Other (30%)
- Germany had the highest number of patients

### 🔎 Diagnostic
- Stage 0 patients had the highest survival years; Stage IV had the lowest
- High-risk patients (severity > 7) = 44 out of 531 (8.3%)
- Treatment cost showed an increasing trend from 2022 to 2024

### 🔮 Predictive
- High Genetic_Risk + Smoking + Obesity = Higher severity scores
- Stage IV patients predicted to have above-average treatment costs
- Treatment costs will continue rising year over year

### 💊 Prescriptive
- Prioritise early screening (Stage 0/I) — significantly higher survival rates
- Target lifestyle risk factors through public health campaigns
- Increase cancer awareness programs in high-patient countries

---

## 🏁 Conclusion

This project successfully demonstrated an end-to-end data analysis workflow using Microsoft Excel and Power BI. The raw cancer patient dataset was systematically cleaned using 12 data cleaning steps and visualized through 8 meaningful charts in an interactive Power BI dashboard with 5 DAX measures and 3 slicers for dynamic filtering.

---

## 👩‍💻 Author

MONIKA.S 
ANB18
