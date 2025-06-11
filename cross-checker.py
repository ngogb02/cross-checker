from openpyxl import load_workbook
from pydantic import BaseModel
import pandas as pd

# 1st --------------------------------------[Clean up markup - remove text with strike-through]-----------------------------------------------

# Load the mark up sheet into the Workbook object from openpyxl.load_workbook
workbook = load_workbook("sample_data/SRO658Markup.xlsx")

# Always working with the first worksheet
worksheet = workbook.worksheets[0]

for row in worksheet.iter_rows(min_row=2): # 1 = header, 2 = first data row (skipping the header row)
    for cell in row:
        # Only care about cells with string values that are rich text
        if not cell.value or isinstance(cell.value, str):
            continue # continue to next iteration - doesn't proceed with next if-condition
        
        # If entire cell is strike-through, clear the cell
        if cell.font.strike:
            cell.value = ""
            continue # continue to next interation - doesn't proceed with next if-condition
        
        # Fetch the attr 'rich_text' if it exist in cell, otherwise return None
        if getattr(cell, 'rich_text', None):
            new_text = []
            # A “run” = one contiguous slice of characters that all share identical formatting (font name, size, bold, strike, color …).
            # cell.rich_text is a tuple of run objects with attr "text" and "font".
            for run in cell.rich_text: 
                if not run.font.strike:
                    new_text += run.text 
            cell.value = new_text


# Read and load the markup into dataframe.
df_markup = pd.read_excel("sample_data/SRO658Markup.xlsx", engine="openpyxl")
# Read and load the export from DOORS into dataframe.
df_doors = pd.read_excel("sample_data/SRO658_DOORS.xlsx", engine="openpyxl")

