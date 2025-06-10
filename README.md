
# FastAPI Excel Processor

This FastAPI application processes an Excel sheet (`capbudg.xls`) and exposes RESTful endpoints to interact with the data.

## File Structure
```
.

├── main.py
├── Data/
│   └── capbudg.xls
├── requirements.txt
└── README.md
```

## Features

- List all tables (Excel sheet names)
- View row names from a specific table
- Calculate sum of numerical values in a given row

---

## Setup Instructions

1. **Clone the repository**

```bash
git clone https://github.com/KaranMotwani22/IRIS_Assignment_1
cd IRIS_Assignment_1
```

2. **Install dependencies**

```bash
pip install -r requirements.txt
```

3. **Run the server**

```bash
uvicorn main:app --reload --port 9090
```

---

## API Endpoints

### 1. `GET /list_tables`
Returns a list of all tables in the Excel sheet.

**Response:**
```json
{
  "tables": ["INITIAL INVESTMENT","CASHFLOW DETAILS","DISCOUNT RATE"]
}
```

---

### 2. `GET /get_table_details?table_name=CASHFLOW DETAILS`
Returns first-column values (row names) from a specific table.

**Response:**
```json
{
  "table_name": "CASHFLOW DETAILS",
  "row_names": [
    "Revenues in  year 1=",
    "Var. Expenses as % of Rev=",
    "Fixed expenses in year 1=",
    "Tax rate on net income=",
    "If you do not have the breakdown of fixed and variable",
    "expenses, input the entire expense as a % of revenues."
  ]
}
```

---

### 3. `GET /row_sum?table_name=INITIAL INVESTMENT&row_name=Initial Investment=`
Returns sum of numerical values in a specific row.

**Response:**
```json
{
  "table_name": "INITIAL INVESTMENT",
  "row_name": "Initial Investment=",
  "sum": 50000
}
```

---

## Testing with Postman

### Base URL: `http://localhost:9090`

You can import the `postman_collection.json` file into Postman to test the API.

---

## My Insights

### Potential Improvements
- The excel file provided should have separated sheets for different tables which can help making the code more robust and reducing potential errors.
- Support for `.xlsx` and `.csv` formats.
- File upload endpoint for dynamic file processing.
- UI dashboard to visualize Excel data.

### Missed Edge Cases
- The Table names are recognized based on Italic font and borders for table, so multiple tables can be recognized in the same sheet, if you don't make the table-name italic it will be missed.
- As I am using a dictionary for storing tables duplicate names can cause issues.
- Empty Excel files or sheets.
- Non-numeric or malformed values in data rows.

## Logic

1. I've made the table-names italic so that they can be recognized in code and covered all the table structures with a border to recognize the no. of rows and column of the given table
2. The tables are then stored in a dictionary named tables with their names as keys and the data as values in the form of dataframes. 

---

##  Author

Karan Motwani

---
