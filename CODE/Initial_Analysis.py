import util
import pandas as pd

"""
Prompts User to Submit Location of Excel File.
Checks to ensure Excel File.
"""
def prompt_excel_file():
    while True:
        file_path = input("Please provide the path to your Excel file: ")
        
        # Check if the file exists
        if not os.path.exists(file_path):
            print("Error: The file does not exist. Please try again.")
            continue
        
        # Check if the file is an Excel file
        if not file_path.lower().endswith(('.xlsx', '.xls')):
            print("Error: The file is not an Excel file. Please provide a valid Excel file (.xlsx or .xls).")
            continue
        
        try:
            # Attempt to load the Excel file
            df = pd.read_excel(file_path)
            print(f"Success: Loaded {file_path}.")
            return file_path
        except Exception as e:
            print(f"Error: Unable to read the Excel file. Details: {e}")
            continue

# Example usage
if __name__ == "__main__":
    path_raw_xl = prompt_excel_file()
    
    path_export = r'V0 EXAMPLE'
    """
    All calculations based off of this Database
    """
    util.generate_ADE_DB(path_raw_xl, path_export)
    """
    Used Pandas DF to Filter for only Neccessary Columns and calculates Current FTE grouping by worker
    Exports it to an excel file with some 
    """
    util.export_xl(path_ADE_db, path_export)

    util.export_ADS(path_ADE_db, path_export_ads)

    #TODO: add prompt for initial run an unassigned allocation budgets
    util.add_buckets(path_export, path_ADE_db)
    util.add_buckets_ads(path_export_ads, path_ADE_db)
    #util.combine_excel_files(path_export, path_export_ads, path_raw_xl, path_final)

    print('Workday Report R0134 converted into FY 25 Salary Allocation.xslx')