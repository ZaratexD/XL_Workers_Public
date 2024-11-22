# XL_Workers_Public
A public version describing the transformation of raw Workday Data into Excel Dashboard. Used for Fiscal analysis of worker salary and allocation.

WorkFlow Pipeline:

1.) Have Excel Report R0314 Employee Details By Organization (Needs security clearence to run via Workday)
  a.) (OPTIONAL) If report has been run before input Allocation_Matrix.xslx, mapping Budget Tags to Budget Name. Budget Name can be custom
2.) Run Initial_Analysis.py
3.) Console will show tags that are currently empty. Input name or leave blank for default name (Budget_1, Budget_2,... Budget_X) etc.
  NOTE: if there are redundant budget tags mapped to different names Allocation Matrix will show Salary_Distributions > 100%
4.) Output: FY 25 Salary Allocation.xslx see below for example and column definitions.

