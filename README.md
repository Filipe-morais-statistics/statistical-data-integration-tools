# statistical-data-integration-tools
This repository contains a set of Python and VBA scripts developed to support the automated integration, validation and transformation of statistical and administrative data used in the compilation of official statistics.
The tools were designed to reduce manual work, improve data quality and ensure consistent ingestion of heterogeneous Excel-based sources into structured statistical databases.
________________________________________
📌 Overview
The scripts in this repository perform three main functions:
1. Automated ingestion of Excel-based data sources (Python)
The Python files (main_DDE.py and main_cga.py) read raw Excel files from external providers, transform the data into harmonised structures, and export ready-to-load datasets for compilation systems.
2. Data cleaning, validation and structure harmonisation
The processes include:
•	time-series reconstruction
•	recoding and harmonisation of instrument classifications
•	detection and correction of formatting anomalies
•	integration of multiple sheets and heterogeneous formats
•	preparation of period identifiers and metadata
•	reconciliation with auxiliary datasets
3. VBA automation for reporting processes
The VBA module automates the transformation of Excel inputs into the exact structured format required for ingestion into reporting databases, eliminating repetitive manual formatting.
________________________________________
📁 Files in this repository
📌 main_DDE.py — Source data 1 Integration
This script:
•	loads multiple Excel sheets from the data source 1
•	extracts table structures using the tabulizer package
•	cleans and harmonises month/year formats
•	standardises instrument classifications
•	builds a structured dataset with variables such as:
o	Period, Currency, Instrument, Maturity, Instrument Detail, Value (MEUR)
•	exports the harmonised table for ingestion into the central statistical repository
Used for: official monthly public debt compilation.
________________________________________
📌 main_cga.py — Source data 2 Portfolio Integration
This script:
•	reads the monthly securities portfolio file of the Source data 2
•	cleans structure, removes totals and irrelevant lines
•	reshapes data from wide to long (nominal value vs market value)
•	creates the Period variable based on the reporting date
•	prepares a load-ready dataset with:
o	Period, Security, Metric Type, Value
Used for: official monthly public debt compilation.
________________________________________
📌 VBA-Module1.bas — Excel → Reporting Database Transformation
The VBA module:
•	converts raw Excel information into a clean, standardised table
•	enforces column ordering, naming conventions and validation rules
•	prepares files automatically for statistical reporting systems
Used for: pre-processing and transformation of reporting data before integration.
________________________________________
🧰 Technologies
•	Python (pandas, numpy, openpyxl, tabulizer)
•	VBA (Excel automation)
•	SQL (downstream integration)
________________________________________
🎯 Purpose
These tools support statistical production by:
•	improving efficiency
•	reducing operational risk
•	reinforcing consistency across datasets
•	automating repetitive data-intensive tasks
They reflect hands-on experience in financial and government finance statistics, data quality assessment and process automation.

