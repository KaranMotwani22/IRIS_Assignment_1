from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import xlrd
from typing import Dict, List
import re

app = FastAPI()
file_path = "./Data/capbudg.xls"  

# Allow all origins for simplicity
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

def clean_value(value):
    #Clean strings like $50,000 or 10% for numeric conversion.
    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)  # Return numbers directly
    if not isinstance(value, str) or not value.strip():
        return None
    # Remove $, ,, and whitespace; handle %
    cleaned = re.sub(r'[\$,]', '', value.strip())
    if cleaned.endswith('%'):
        cleaned = cleaned.rstrip('%')
        try:
            return float(cleaned) #/ 100  
        except ValueError:
            return None
    try:
        return float(cleaned)
    except ValueError:
        return None

def load_excel_sheets() -> Dict[str, pd.DataFrame]:
    try:
        workbook = xlrd.open_workbook(file_path, formatting_info=True)
        sheet = workbook.sheet_by_index(0)  
        font_list = workbook.font_list

        # Dictionary to store table names and their DataFrames
        tables = {}

        # List to store table metadata (name, row, start_col, end_col)
        table_info = []

        # Step 1: Identify headers and their column ranges based on black outlines
        for row in range(sheet.nrows):
            current_table_name = None
            start_col = None
            for col in range(sheet.ncols):
                cell_value = sheet.cell_value(row, col)
                cell_xf_index = sheet.cell_xf_index(row, col)
                cell_xf = workbook.xf_list[cell_xf_index]
                font = font_list[cell_xf.font_index]
                borders = cell_xf.border

                # Check for italicized header to identify potential table start
                if cell_value and isinstance(cell_value, str) and len(cell_value.strip()) > 5 and font.italic:
                    current_table_name = cell_value.strip().replace('\n', ' ')
                    start_col = col
                    # Look for black outline to determine end_col
                    end_col = col
                    while end_col < sheet.ncols:
                        try:
                            next_cell_xf_index = sheet.cell_xf_index(row, end_col)
                            next_cell_xf = workbook.xf_list[next_cell_xf_index]
                            next_borders = next_cell_xf.border
                            # Check for black outline (border style > 0 indicates a visible border)
                            # Assuming black outline means at least one border (top, bottom, left, right) is present
                            has_border = (
                                next_borders.top_line_style > 0 or
                                next_borders.bottom_line_style > 0 or
                                next_borders.left_line_style > 0 or
                                next_borders.right_line_style > 0
                            )
                            if not has_border and end_col > start_col:
                                break  # End of table when no border is found
                            end_col += 1
                        except IndexError:
                            break
                    if end_col > start_col:  # Valid table found
                        table_info.append((current_table_name, row, start_col, end_col))
                    current_table_name = None
                    start_col = None

        # Step 2: Extract table data
        for table_name, header_row, start_col, end_col in table_info:
            table_data = []
            expected_col_count = end_col - start_col  # Number of columns based on outline

            # Collect data from rows below header
            for row_idx, row in enumerate(range(header_row + 1, sheet.nrows)):
                # Stop if we hit another italicized header
                stop = False
                for col in range(sheet.ncols):
                    cell_value = sheet.cell_value(row, col)
                    if cell_value and isinstance(cell_value, str) and len(cell_value.strip()) > 5:
                        cell_xf_index = sheet.cell_xf_index(row, col)
                        cell_xf = workbook.xf_list[cell_xf_index]
                        font = font_list[cell_xf.font_index]
                        if font.italic:
                            stop = True
                            break
                if stop:
                    break

                # Determine the starting column for this row
                current_start_col = start_col
                if row_idx > 0:  # Check left columns for non-empty cells (from second data row)
                    left_col = start_col - 1
                    while left_col >= 0:
                        try:
                            left_value = sheet.cell_value(row, left_col)
                            if left_value and (isinstance(left_value, (int, float)) or (isinstance(left_value, str) and left_value.strip())):
                                current_start_col = left_col
                            else:
                                break
                        except IndexError:
                            break
                        left_col -= 1

                # Collect row data from current_start_col to end_col
                row_data = []
                has_data = False
                for col in range(current_start_col, end_col):
                    try:
                        cell_value = sheet.cell_value(row, col)
                        cell_type = sheet.cell_type(row, col)

                        # For first column, handle percentage values
                        # if col == current_start_col:
                        if cell_type == xlrd.XL_CELL_NUMBER and 0 < cell_value <= 1:
                            # Heuristic: assume small numbers (0 < x <= 1) are percentages
                            display_value = f"{cell_value * 100:.2f}%"
                            row_data.append(display_value)
                        else:
                            row_data.append(str(cell_value) if cell_value else '')
                        # else:
                        #     row_data.append(cell_value)
                        if cell_value and (isinstance(cell_value, (int, float)) or (isinstance(cell_value, str) and cell_value.strip())):
                            has_data = True
                    except IndexError:
                        row_data.append(None)

                # Check column count
                if has_data:
                    if len(row_data) >= expected_col_count:
                        table_data.append(row_data[:expected_col_count])  # Trim to expected columns
                elif not has_data and row_data and all(v is None or (isinstance(v, str) and not v.strip()) for v in row_data):
                    break  # Stop at empty row

            # Convert to DataFrame
            if table_data:
                df = pd.DataFrame(table_data)
                if not df.empty and df.shape[1] > 0:
                    tables[table_name] = df

        print("Table Names and Column Counts found in the first sheet:")
        for name, _, start_col, end_col in table_info:
            print(f"{name}: {end_col - start_col} columns (from col {start_col} to {end_col - 1})")

        return tables
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error reading Excel file: {str(e)}")

@app.get("/")
def root():
    return {"message": "FastAPI Excel Processor is running!"}

@app.get("/list_tables")
def list_tables():
    tables = load_excel_sheets()
    return {"tables": list(tables.keys())}

@app.get("/get_table_details")
def get_table_details(table_name: str = Query(...)):
    tables = load_excel_sheets()
    if table_name not in tables:
        raise HTTPException(status_code=404, detail=f"Table '{table_name}' not found. Available tables: {list(tables.keys())}")
    df = tables[table_name]
    # Assume first column contains row names, drop NaN and convert to strings
    row_names = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    return {"table_name": table_name, "row_names": row_names}

@app.get("/row_sum")
def row_sum(table_name: str = Query(...), row_name: str = Query(...)):
    tables = load_excel_sheets()
    if table_name not in tables:
        raise HTTPException(status_code=404, detail=f"Table '{table_name}' not found. Available tables: {list(tables.keys())}")
    df = tables[table_name]
    # Find the row matching row_name in the first column
    for i, value in enumerate(df.iloc[:, 0].astype(str)):
        if value.strip() == row_name.strip():
            # Clean and convert values in remaining columns
            cleaned_values = [clean_value(val) for val in df.iloc[i, 1:]]
            # Filter out None and sum
            numeric_values = [v for v in cleaned_values if v is not None and not pd.isna(v)]
            if not numeric_values:
                return {
                    "table_name": table_name,
                    "row_name": row_name,
                    "sum": 0.0
                }
            return {
                "table_name": table_name,
                "row_name": row_name,
                "sum": float(sum(numeric_values))
            }
    raise HTTPException(status_code=404, detail=f"Row name '{row_name}' not found in table '{table_name}'")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)