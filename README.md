# Coca-Cola (KO) Financial Analysis

## What This Project Does
This Python script pulls live Coca-Cola financial data from Yahoo Finance
and automatically calculates key FP&A and equity research metrics,
then exports a formatted Excel report.

## What It Analyzes
- Income Statement (Revenue, Gross Profit, Operating Income, Net Income)
- Gross, Operating, and Net Margins (2022-2025)
- Free Cash Flow = Operating Cash Flow - Capital Expenditures
- Year over Year Revenue Growth Rates
- Last 5 days of KO stock price
- KO vs PEP comparable margin analysis

## Key Findings
- KO gross margin expanded from 58.1% in 2022 to 61.6% in 2025
- Operating margin improved from 28.0% to 31.1% over the same period
- FCF dropped significantly from $9.7B in 2023 to $4.7B in 2024
- Revenue growth is decelerating: 6.4% in 2023 → 1.9% in 2025
- KO operating margin of 31.1% is nearly double PEP's 14.4%

## Libraries Used
- pandas — data manipulation and analysis
- yfinance — pulls live financial data from Yahoo Finance
- openpyxl — exports formatted Excel report

## Output
Generates KO_Analysis.xlsx with three formatted sections:
income statement, margin analysis, and free cash flow

## Author
Christopher Moya Ramirez
Finance Major, Mercy University