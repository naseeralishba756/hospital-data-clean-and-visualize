# hospital-data-clean-and-visualize
# Hospital Finance Data Analysis - Top 5 Sales Days

## 📊 Project Overview
This project is an end-to-end analysis of a Hospital's Finance General Ledger containing 12,003+ journal entries. The main objective was to identify the top-performing sales days for **Account 4010 (Sales Revenue)** using both Excel and Python.

## 📁 Dataset
- **File:** `PNRao_Finance_General_Ledger_Tr.xlsx` / `Hospital data.xlsx`
- **Source:** Hospital Finance System (General Ledger - Detail)
- **Total Records:** 12,003 entries
- **Date Range:** 2023 - 2024
- **Key Columns Used:** Date, Account, Description, Credit, Debit, Balance

## 🛠️ Tools & Technologies Used
- **Excel / WPS Office:** Data Filtering, Pivot Table, Pivot Chart
- **Python (Google Colab):** Data Analysis & Automation
- **Libraries:** `pandas`, `matplotlib.pyplot`, `openpyxl`

## 🔍 Analysis Workflow

### 1. Excel Analysis (WPS Office)
1.  Filtered the dataset for `Account = 4010`.
2.  Created a Pivot Table: `Rows -> Date`, `Values -> SUM of Credit`.
3.  Sorted descending and extracted Top 5 days.
4.  Created a Pivot Bar Chart for visualization.

### 2. Python Analysis (Google Colab)
```python
import pandas as pd
import matplotlib.pyplot as plt

# Load data
df = pd.read_excel("Hospital data.xlsx")

# Filter for Account 4010
sales = df[df['Account'] == 4010]

# Get Top 5 Sales Days
top5 = sales.groupby('Date')['Credit'].sum().sort_values(ascending=False).head(5)

# Visualization
top5.plot(kind='bar', figsize=(10,5), color='#4CAF50')
plt.title('Top 5 Sales Days - Account 4010')
plt.ylabel('Total Sales ($)')
plt.xticks(rotation=45)
plt.show()