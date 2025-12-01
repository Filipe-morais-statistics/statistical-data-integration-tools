# statistical-data-integration-tools
This repository contains a set of Python and VBA scripts developed to support the automated integration, validation and transformation of statistical and administrative data used in the compilation of official statistics.
The tools were designed to reduce manual work, improve data quality and ensure consistent ingestion of heterogeneous sources based in Excel into structured statistical databases.

**Overview**

The scripts in this repository perform three main functions:
1. Automated ingestion of Excel-based data sources (Python). The Python files (main_DDE.py and main_cga.py) read raw Excel files from external providers, transform the data into harmonised structures, and export ready-to-load datasets for compilation systems.
2. Data cleaning, validation and structure harmonisation. The processes include:
- time-series reconstruction;
-	recoding and harmonisation of instrument classifications;
-	detection and correction of formatting anomalies;
-	integration of multiple sheets and heterogeneous formats;
-	preparation of period identifiers and metadata;
-	reconciliation with auxiliary datasets.
3. VBA automation for reporting processes. The VBA module automates the transformation of Excel inputs into the exact structured format required for ingestion into reporting databases, eliminating repetitive manual formatting.

**Files in this repository**

A. main_DDE.py — "Data source 1" integration. This script :
-	loads multiple Excel sheets from the "data source 1";
-	extracts table structures using the tabulizer package;
-	cleans and harmonises month/year formats;
-	standardises instrument classifications;
-	builds a structured dataset with variables (Period, Currency, Instrument, Maturity, Instrument Detail, Value);
-	exports the harmonised table for ingestion into the central statistical repository.
Used for: monthly public debt compilation.

B. main_cga.py — "Data source 2" integration. This script:
-	reads the monthly securities portfolio file of the "data source 2";
-	cleans structure, removes totals and irrelevant lines;
-	reshapes data from wide to long (nominal value vs market value);
-	creates the Period variable based on the reporting date;
-	builds a structured dataset with variables (Period, Security, Metric Type, Value).
Used for: monthly public debt compilation.

C. VBA-Module1.bas — Transformation of Excel into a reporting file. The VBA module:
-	converts raw Excel information into a clean, standardised table;
-	enforces column ordering, naming conventions and validation rules;
-	prepares files automatically for statistical reporting systems.
Used for: pre-processing and transformation of reporting data before transmissions.

**Programming tools used:**

-	Python (pandas, numpy, openpyxl, tabulizer);
-	VBA (Excel automation).
