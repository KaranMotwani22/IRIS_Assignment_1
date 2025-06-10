
# FastAPI Excel Processor

This FastAPI application processes an Excel sheet (`capbudg.xls`) and exposes RESTful endpoints to interact with the data.

## 📂 File Structure
```
.

├── main.py
├── Data/
│   └── capbudg.xls
├── requirements.txt
└── README.md
```

## 🚀 Features

- ✅ List all tables (Excel sheet names)
- ✅ View row names from a specific table
- ✅ Calculate sum of numerical values in a given row

---

## 🔧 Setup Instructions

1. **Clone the repository**

```bash
git clone https://github.com/your-username/excel-api.git
cd excel-api
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

## 📬 API Endpoints

### 1. `GET /list_tables`
Returns a list of all tables in the Excel sheet.

**Response:**
```json
{
  "tables": ["Initial Investment", "Revenue Projections", "Operating Expenses"]
}
```

---

### 2. `GET /get_table_details?table_name=Initial Investment`
Returns first-column values (row names) from a specific table.

**Response:**
```json
{
  "table_name": "Initial Investment",
  "row_names": [
    "Initial Investment=",
    "Opportunity cost (if any)=",
    ...
  ]
}
```

---

### 3. `GET /row_sum?table_name=Initial Investment&row_name=Tax Credit (if any )=`
Returns sum of numerical values in a specific row.

**Response:**
```json
{
  "table_name": "Initial Investment",
  "row_name": "Tax Credit (if any )=",
  "sum": 10
}
```

---

## 🧪 Testing with Postman

### Base URL: `http://localhost:9090`

You can import the `postman_collection.json` file into Postman to test the API.

---

## 🧠 Your Insights

### 🔁 Potential Improvements
- The excel file provided should have separated sheets for different tables which can help making the code more robust and reduced potential errors.
- Support for `.xlsx` and `.csv` formats.
- File upload endpoint for dynamic file processing.
- UI dashboard to visualize Excel data.

### ⚠️ Missed Edge Cases
- The Table names are recognized based on Italic font and borders for table, so multiple tables can be recognized in the same sheet, if you don't make the table-name italic it will be missed.
- As I am using a dictionary for storing tables duplicate names can cause issues.
- Empty Excel files or sheets.
- Non-numeric or malformed values in data rows.

---

## 👨‍💻 Author

Karan Motwani

---
