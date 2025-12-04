# Health Care Data Analysis (Interactive Dashboard creation using MS Excel)

# Introduction

This health Dashboard was created in Microsoft Excel. The dataset includes information such as patient visits, diagnoses, doctors, and treatment costs, but it contains inconsistencies and errors that hinder analysis.

# Project Objectives

- To Clean and organize raw health data to ensure accuracy and consistency.
- To Apply appropriate Excel formatting techniques to enhance data readability.
- To Use Excel formulas from multiple categories (Logical, Statistical, Lookup, Math & Date) to perform meaningful calculations and analysis.
- To Create PivotTables and PivotCharts to summarize and visualize key health metrics.
- To Present findings in a clear, professional report that communicates insights effectively.

# Dataset Used
- <a href= https://github.com/Onyinyechukwu5/Health-Care-Data-Analysis/blob/main/Health_Dirty.csv> Dataset </a>

# Questions

## Part 1: Data Cleaning
- Remove any duplicate records based on PatientID.
- Convert VisitDate into a proper Excel date format (dd-mmm-yyyy).
- Handle missing data: Replace blank Diagnosis with 'Unknown'; Replace blank
Doctor with 'Unassigned'.
- Remove the empty column AGE RANGE (you will rebuild this later).
- Standardise text values: Ensure Gender only has Male, Female, Other; Remove extra spaces from
Doctor, Treatment, and Diagnosis fields.

## Part 2: Formatting
- Apply bold text + background colour to column headers.
- Format Cost as currency with 2 decimal places.
- Format VisitDate as dd-mmm-yyyy.
- Add conditional formatting: Highlight patients older than 60; Highlight treatments
costing more than 3000

## Part 3: Formulas
- Create a new sheet called Formulas and answer the following:
- Logical Functions: Use IF to create a new column 'Age Group'.
- Statistical Functions: Find average, max, min of Cost; Count patients with
Diabetes.
- Lookup & Reference: Create lookup table with Treatment Types and Standard
Costs; use VLOOKUP.
- Math & Trig: Add a column 'Cost with Tax' = Cost × 1.1; Use ROUND, ROUNDUP,
ROUNDDOWN.
- Date Functions: Extract Year and Month from VisitDate; Rebuild AGE RANGE with IF

## Part 4: Pivot Tables & Charts
- Create a new sheet called Analysis.
- PivotTable: Total Cost by Doctor.
- PivotTable: Count of Patients by Gender.
- PivotTable: Average Cost by Diagnosis.
- PivotTable: Age Group Distribution.
- Insert PivotCharts: Pie chart (Gender distribution), Column chart (Total Cost by Diagnosis)

# Analysis and Interactive Dashboard
- <a href = https://github.com/Onyinyechukwu5/Health-Care-Data-Analysis/blob/main/Health_Data%20Analysis.xlsx> Interactive Dashboard <a/>
# Process

## Data Cleaning
Removed duplicate records from patientID. To ensure each patient record is unique.
Converted VisitDate into a Proper Excel Date Format (dd-mmm-yyyy). To ensure all visit dates are in a valid and readable date format.
  
# Handling Missing Data
Replaced blank fields with meaningful placeholder values.
For Diagnosis column missing data was replaced with “unknown” following these steps;
For the Doctor column missing data was replaced with “Unassigned” following the same process.
This helps to keep data consistent and prevents formula or pivot table errors caused by empty cells.

# Standardizing Text Values
To Ensure all text data follows a consistent format for accuracy in analysis.
For Gender column:
I Replaced; M or male → Male, Replaced; F or female → Female and Replaced any unclear entries with Other.

# Data Formatting
Ensure all cost values are properly formatted as currency for consistency.
Cost column was formatted to currency and also to 2 decimal places.

# Add Conditional Formatting
Conditional formatting was don on the Age column to Highlight Patients Older Than 60

# Formulas and Calculations
Logical Functions – To Create ‘Age Group’
Patients was categorised into age groups using an IF function.
=IF(C2<12,"Child",IF(C2<19,"Teenager", IF(C2<59,"Adult","Old")))
This formula checks the patient’s age and assigns the correct category.

# Statistical Functions
To summarize treatment costs. Find the Average, Maximum, and Minimum of Cost
Formulas used:
Since the Cost is in column H then; =AVERAGE(H2:H100)
=MAX(H2:H100)
=MIN(H2:H100)

# Count Patients with Diabetes
To get how many patients have “Diabetes” COUNTIF formula was used and since diagnosis is in column E we have; =COUNTIF(E2:E100,"Diabetes")

# Lookup & Reference – VLOOKUP for Standard Costs
Used a lookup table to find the Standard Cost for each treatment type.
Formula:
 =VLOOKUP(Treatment, TreatmentLookup, 2, FALSE)
This retrieves the standard cost based on the treatment type.

# Dashboard
<img width="819" height="1080" alt="image" src="https://github.com/user-attachments/assets/e64bae2d-94be-428b-9bb5-19f1f8f490c7" />

<img width="896" height="1080" alt="image" src="https://github.com/user-attachments/assets/fb4e0778-f432-4ea0-b747-502965defc74" />

<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/4b5064a7-2eb6-402a-b747-80ce2ce4428d" />

<img width="1777" height="790" alt="Dashboard" src="https://github.com/user-attachments/assets/265617dd-7fef-4063-b418-6662f2613809" />

PROJECT INSIGHT
- Total Cost by Doctor
Five doctors are handling most of the cost equally (≈ $55M each).
One category “Not Assigned” shows a significantly lower cost (8.58M), suggesting incomplete or missing doctor assignments.

- Patient Distribution by Gender
Gender proportion is well-balanced: Male: 34% Female: 33% Other: 33%
This indicates no noticeable gender bias in service access.

- Patients by Treatment Type
Surgery is the most utilized treatment type (4,981 patients), closely followed by Therapy and Observation.
Medication has the lowest count (4,971), though still close to the others.
Overall, the treatment distribution is nearly uniform, but surgery slightly leads.

- Monthly Patient Visits
Visits fluctuate slightly between 8,700 – 9,700 monthly.
February shows the lowest visits (8,792), while March, August, and December show peaks.
No major seasonal dip except February.

- Cost per Diagnosis
Costs across diagnoses are very similar, averaging around $2.50K per diagnosis.
Migraine has the lowest cost (2.48K), while Flu, Cancer, and Asthma share the highest (2.51K).
This suggests standardized pricing or similar treatment intensity.

- Age Distribution
Adults (27%) and Old patients (27%) form the largest groups.
Teenagers (21%) and Children (25%) are also significant.
This shows the facility serves a broad age range with slight emphasis on Adult and Old demographics.

- Number of Patients
The dashboard shows 5,000 patients in total.
Average cost per treatment is $2.50K, aligning with cost-by-diagnosis data.

# RECOMMENDATION
- Investigate “Not Assigned” Doctor Records
The low cost suggests incomplete data entry.
Recommend enforcing mandatory doctor assignment to improve reporting accuracy.

- Optimize Resource Allocation for Surgery
Since surgery has the highest patient volume, ensure adequate staffing, equipment, and OR time.

- Age-Specific Programs

With high Adult and Old patient numbers, consider introducing: Chronic disease management programs and Preventive check-ups for older adults.
Meanwhile, maintain pediatric and adolescent service quality.

-Strengthen Gender-Specific Health Awareness
With almost equal distribution across genders, consider targeted programs: Women's health screenings, Men's wellness checks and Inclusivity programs for “Other” gender category.
