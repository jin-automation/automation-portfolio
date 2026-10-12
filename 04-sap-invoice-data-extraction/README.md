# SAP Invoice Data Extraction Automation

**SAP ERP/ECC | Excel VBA | SAP GUI Scripting | Order Management**

Automated invoice data retrieval for multiple Delivery Numbers in SAP ECC, reducing a **20–30 minute daily process to under 5 minutes**.

---

## Business Problem

Invoice details had to be manually traced from individual Delivery Numbers through SAP Document Flow and then copied into Excel.

For larger delivery lists, repeating this process for each record required multiple SAP lookups and manual Excel updates.

---

## Solution

I developed an Excel VBA automation using SAP GUI Scripting that:

1. Reads Delivery Numbers from Excel.
2. Opens each Delivery in SAP using **VL03N**.
3. Navigates through **SAP Document Flow** to locate the related Invoice.
4. Retrieves the **Invoice Number, Invoice Date, and Net Value**.
5. Updates the results directly in Excel.
6. Automatically continues through the remaining Delivery Numbers.

### Information Retrieved

- **Invoice Number**
- **Invoice Date**
- **Net Value**

### Data Flow

**Excel Delivery Number → SAP VL03N → Document Flow → Invoice → Excel Update**

---

## Business Impact

- Reduced daily invoice processing time from **20–30 minutes to under 5 minutes**.
- Automated repetitive invoice tracing through SAP Document Flow.
- Reduced manual SAP lookups and Excel data entry.

---

## Demo

![SAP Invoice Data Extraction Demo](./image/excel-demo.png)

[View Demo Workbook](./demo/SAP_invoice_data_extraction_demo.xlsx)

[View VBA Source Code](./src/SAP_invoice_data_extraction.bas)

---

## Technologies

**SAP ERP/ECC | SAP GUI Scripting | Excel VBA | Order Management | Process Automation**

---

*Portfolio version — company, customer, delivery, invoice, and production data have been removed or replaced with fictional examples.*
