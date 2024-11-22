# XL_Workers_Public

A public version demonstrating the transformation of raw Workday data into an Excel dashboard. This tool is used for fiscal analysis of worker salary and allocation.

---

## Workflow Pipeline

### Step 1: Obtain the Report
- Run the **R0314 Employee Details By Organization** report in Workday.  
  *(Note: Security clearance is required to access this report.)*

- **Optional:**  
  If the report has been run previously, input the `Allocation_Matrix.xlsx` file. This file maps **Budget Tags** to **Budget Names**.  
  - Custom Budget Names can be defined here.

### Step 2: Run the Analysis Script
- Execute `Initial_Analysis.py`.

### Step 3: Review Console Outputs
- The script will display any **empty tags**.  
  - Input names for these tags or leave them blank to use default names (`Budget_1`, `Budget_2`, ..., `Budget_X`, etc.).  
  - **Note:** If redundant budget tags are mapped to different names, the **Allocation Matrix** will show **Salary Distributions > 100%**.

### Step 4: View the Output
- The script generates an Excel file: **`FY 25 Salary Allocation.xlsx`**.  
  - See below for an example and column definitions.

---

## Example Output & Column Definitions
*(TODO)*

---

## Notes
- Ensure security clearance for Workday access.
- Custom Budget Names help improve clarity in fiscal analysis.
- Redundant mappings should be resolved to avoid over-allocation warnings.

---
