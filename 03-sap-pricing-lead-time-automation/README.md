# SAP Pricing & Lead Time Automation

**SAP ERP/ECC | Excel VBA | SAP GUI Scripting | Order Management**

Automated pricing and lead time checks for multiple materials in SAP ECC, retrieving key information directly into Excel and reducing repetitive manual lookups.

---

## Business Problem

Pricing and delivery information had to be checked individually in SAP for multiple materials.

Checking **List Price, Net Price, Discount, and Lead Time** repeatedly was time-consuming, especially when processing larger material lists.

---

## Solution

I developed an **Excel VBA automation using SAP GUI Scripting** that:

1. Reads material and order information from Excel.
2. Retrieves pricing and delivery information from SAP.
3. Updates the results automatically in Excel.
4. Continues through the remaining materials.

### Information Retrieved

- **List Price**
- **Net Price**
- **Discount**
- **Plant**
- **Estimated Delivery Date / Week**

### Data Flow

**Excel Input → SAP ECC → Pricing & Lead Time Lookup → Excel Update**

---

## Business Impact

- Reduced repetitive SAP pricing and availability checks.
- Processed multiple materials automatically.
- Reduced manual Excel entry.
- Improved efficiency when checking pricing and delivery information.

---

## Demo

![SAP Pricing & Lead Time Automation Demo](./images/excel-demo.png)

[View Demo Workbook](./demo/SAP_pricing_lead_time_demo.xlsx)

[View VBA Source Code](./src/SAP_pricing_lead_time.bas)

---

## Technologies

**SAP ERP/ECC | SAP GUI Scripting | Excel VBA | Process Automation**

---

*Portfolio version — company, customer, pricing, material, and production data have been removed or replaced with fictional examples.*
