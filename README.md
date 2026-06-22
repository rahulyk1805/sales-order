# 📦 Pending Orders Processor

A Streamlit-based automation tool designed to process and analyze pending sales orders. The application accepts raw sales order exports or pre-processed order data and automatically generates a structured, business-ready Excel workbook containing production planning, order prioritization, external order tracking, and raw data analysis.

The tool eliminates manual spreadsheet processing, improves planning accuracy, and significantly reduces the time required to prepare pending order reports.

---

# 🚀 Project Overview

The Pending Orders Processor automates the transformation of sales order data into actionable business reports.

It supports two input methods:

## Mode A – Raw JSSOReport Files

* Upload two raw JSSOReport exports.
* Automatically extracts and derives:

  * Brand
  * Depot
  * SKU Information
  * Pending Quantity
  * Order Details
* Processes and structures the data for reporting.

## Mode B – Pre-Built Raw Data File

* Upload a single file containing a **Raw Data** sheet.
* Uses pre-computed fields directly.
* Skips derivation logic and generates reports instantly.

Both modes generate the same standardized output workbook.

---

# ✨ Key Features

## 📊 Automated Data Processing

* Automatic file type detection
* Intelligent data validation
* Automated order categorization
* Quantity and unit handling

## 🏭 Production Planning Support

* Generates production planning sheets
* Helps prioritize manufacturing requirements
* Supports operational decision-making

## 📋 FIFO Order Queue Management

* Creates order queues based on order dates
* Ensures first-in-first-out prioritization
* Improves order fulfillment efficiency

## 🌍 External Order Tracking

* Separates and tracks external orders
* Improves visibility of pending commitments

## 📑 Multi-Sheet Excel Output

Generates a professionally formatted workbook containing:

1. Production Planning
2. Order Queue (FIFO)
3. External Orders
4. Raw Data

## 🎨 User-Friendly Interface

* Modern Streamlit web interface
* Simple file upload workflow
* Download-ready output files
* No coding knowledge required

---

# 🔄 Processing Pipeline

```text
Input Files
      │
      ▼
File Validation
      │
      ▼
Input Mode Detection
(JSSOReport / Raw Data)
      │
      ▼
Data Extraction
      │
      ▼
Business Rule Processing
      │
      ▼
Brand & Depot Mapping
      │
      ▼
Pending Order Analysis
      │
      ▼
FIFO Queue Generation
      │
      ▼
Production Planning Creation
      │
      ▼
Formatted Excel Workbook Output
```

---

# 🏗️ Project Architecture

```text
Pending Orders Processor
│
├── app.py
│   ├── Streamlit UI
│   ├── File Upload Handling
│   ├── User Interaction
│   └── Download Output
│
├── processor.py
│   ├── Data Processing Engine
│   ├── Validation Logic
│   ├── Business Rules
│   ├── Workbook Generation
│   └── Report Formatting
│
├── requirements.txt
│   └── Project Dependencies
│
└── config.toml
    └── Streamlit Configuration
```

---

# 🛠️ Technologies Used

* Python
* Streamlit
* Pandas
* OpenPyXL
* Excel Automation
* Data Processing & Validation

---

# 📈 Business Benefits

### Time Savings

Reduces manual Excel processing from hours to minutes.

### Improved Accuracy

Eliminates human errors in report preparation.

### Standardized Reporting

Ensures consistent output format across all reports.

### Better Production Planning

Provides clear visibility into pending orders and production requirements.

### Faster Decision Making

Converts raw order data into actionable business insights.

### Scalability

Supports different input formats and can be extended for future business requirements.

---

# 🎯 Advantages

* Automated report generation
* Minimal user intervention
* Supports multiple input formats
* Consistent business logic application
* Professional Excel outputs
* Easy deployment through Streamlit
* User-friendly interface
* Reduced operational workload

---

# 🔮 Future Enhancements

* Support for additional order file formats
* Dashboard and analytics integration
* Automated email reporting
* Database integration
* Real-time order monitoring
* Advanced forecasting and planning modules
* Multi-user access control
* ERP integration capabilities

---

# 📦 Installation

```bash
git clone <repository-url>

cd pending-orders-processor

pip install -r requirements.txt

streamlit run app.py
```

---

# ▶️ Usage

1. Launch the Streamlit application.
2. Upload:

   * Two JSSOReport files, or
   * One Raw Data file.
3. Click the processing button.
4. Wait for report generation.
5. Download the generated Excel workbook.

---

# 📄 Output

The generated workbook contains:

* Production Planning Sheet
* Order Queue (FIFO) Sheet
* External Orders Sheet
* Raw Data Sheet

All sheets are automatically formatted and ready for business use.

---

# Live Application
https://sales-order.streamlit.app/
# 👨‍💻 Author

**Rahul Yerunkar**
