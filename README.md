# Health Care Data Analysis (Interactive Dashboard creation using MS Excel)

# Introduction

This project focuses on transforming a raw, unstructured dataset of health records into meaningful insights using Microsoft Excel. The dataset includes information such as patient visits, diagnoses, doctors, and treatment costs, but it contains inconsistencies and errors that hinder analysis.

# Project Objectives

- To Clean and organize raw health data to ensure accuracy and consistency.
- To Apply appropriate Excel formatting techniques to enhance data readability.
- To Use Excel formulas from multiple categories (Logical, Statistical, Lookup, Math & Date) to perform meaningful calculations and analysis.
- To Create PivotTables and PivotCharts to summarize and visualize key health metrics.
- To Present findings in a clear, professional report that communicates insights effectively.

# Dataset Used
- <a href= https://github.com/Onyinyechukwu5/Health-Care-Data-Analysis/commit/0392b9001fb84f0ab18730bec51dba2da295a095> Dataset </a>

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

