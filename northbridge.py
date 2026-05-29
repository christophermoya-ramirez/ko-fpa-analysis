import pandas as pd

print("========================================")
print("   NORTHBRIDGE CONSUMER GOODS — 2025")
print("      FP&A Budget vs. Actual Report")
print("========================================")

# Load all three sheets
variance = pd.read_excel(r"C:\Users\chris\Desktop\python\NorthBridge Consumer Goods.xlsx", sheet_name="VARIANCE REPORT", skiprows=4)
budget = pd.read_excel(r"C:\Users\chris\Desktop\python\NorthBridge Consumer Goods.xlsx", sheet_name="BUDGET", skiprows=1)
actuals = pd.read_excel(r"C:\Users\chris\Desktop\python\NorthBridge Consumer Goods.xlsx", sheet_name="ACTUALS", skiprows=1)

# Clean variance sheet
variance.columns = ["Category", "Budget", "Actual", "$ Variance", "% Variance", "F/U"]
variance = variance[pd.to_numeric(variance["Budget"], errors="coerce").notna()]
variance["Budget"] = variance["Budget"].round(2)
variance["Actual"] = variance["Actual"].round(2)
variance["$ Variance"] = variance["$ Variance"].round(2)

# Clean budget sheet
budget.columns = ["Category", "Jan", "Feb", "Mar", "Apr", "May", "Jun","Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Full Year"]
budget = budget[pd.to_numeric(budget["Jan"], errors="coerce").notna()]

# Clean actuals sheet
actuals.columns = ["Category", "Jan", "Feb", "Mar", "Apr", "May", "Jun","Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Full Year"]
actuals = actuals[pd.to_numeric(actuals["Jan"], errors="coerce").notna()]

# Section 1 - Top variances
print("\n--- SECTION 1: Biggest Misses ---")
unfavorable = variance[variance["F/U"] == "Unfavorable"]
unfavorable = unfavorable.sort_values("$ Variance")
print(unfavorable[["Category", "Budget", "Actual", "$ Variance"]].to_string(index=False))

# Section 2 - Merged budget vs actual
print("\n--- SECTION 2: Full Year Budget vs Actual ---")
merged = pd.merge(budget[["Category", "Full Year"]], actuals[["Category", "Full Year"]], on="Category", suffixes=("_Budget", "_Actual"))
merged["$ Variance"] = (merged["Full Year_Actual"] - merged["Full Year_Budget"]).round(2)
merged["Full Year_Budget"] = merged["Full Year_Budget"].round(2)
merged["Full Year_Actual"] = merged["Full Year_Actual"].round(2)
print(merged.to_string(index=False))

# Section 3 - Quarterly revenue trend
print("\n--- SECTION 3: Actual Revenue by Quarter ---")
actuals["Q1"] = actuals["Jan"] + actuals["Feb"] + actuals["Mar"]
actuals["Q2"] = actuals["Apr"] + actuals["May"] + actuals["Jun"]
actuals["Q3"] = actuals["Jul"] + actuals["Aug"] + actuals["Sep"]
actuals["Q4"] = actuals["Oct"] + actuals["Nov"] + actuals["Dec"]
revenue = actuals[actuals["Category"] == "Revenue"].copy()
revenue["Q1"] = revenue["Q1"].round(2)
revenue["Q2"] = revenue["Q2"].round(2)
revenue["Q3"] = revenue["Q3"].round(2)
revenue["Q4"] = revenue["Q4"].round(2)
print(revenue[["Category", "Q1", "Q2", "Q3", "Q4"]].to_string(index=False))

print("\n========================================")
print("           END OF REPORT")
print("========================================")