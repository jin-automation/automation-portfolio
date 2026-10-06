# SAP Delivery Date Automation

**SAP ERP/ECC | Excel VBA | SAP GUI Scripting | Order Management**

Automated high-volume sales order delivery date changes in SAP ECC, reducing a **40+ minute repetitive manual task to ~10 minutes using 2 Atomation Methods**—by Line Item and by Part Number—with validation controls to reduce update errors.

---

## Business Problem

Delivery date change requests could involve **100+ line items**, requiring each item to be manually checked and updated in SAP.

The manual process took **more than 40 minutes** for a large request and involved repetitive verification of:

- Part Number
- Quantity
- Requested Delivery Date

This created unnecessary manual workload and increased the risk of updating the wrong item.

---

## Solution

I developed two Excel VBA automation methods using SAP GUI Scripting to process high-volume delivery date change requests.

### Method 1 — Search by Line Item

The automation:

1. Reads the Customer PO and Line Item from Excel.
2. Locates the requested line item in SAP ECC.
3. Retrieves the SAP Part Number and Quantity.
4. Validates them against the request data.
5. Updates the delivery date only when both values match.
6. Records the processing result in Excel.

### Method 2 — Search by Part Number

For requests where the Part Number is used as the reference, the automation:

1. Reads the Customer PO and Part Number from Excel.
2. Locates the corresponding material in SAP ECC.
3. Identifies the relevant item.
4. Updates the requested delivery date.
5. Records the processing result in Excel.


### Safety & Validation

> **Part No. and Qty must match SAP.**  
> **Match → Delivery date updated.**  
> **Mismatch → No update.**

This validation prevents the automation from changing the delivery date when the SAP data does not match the request.

---

## Business Impact

| Area | Before | After |
|---|---|---|
| Processing | Manual SAP updates | Automated processing |
| Large request | 40+ minutes | ~10 minutes |
| Validation | Manual checking | Automated Part No. & Qty validation |
| Error control | Dependent on manual verification | Mismatch automatically stops update |
| Result tracking | Manual | Status recorded for each line |

---

## Excel Demo

The demo below uses fictional data to show how customer requests, SAP validation, and processing results are presented.

![SAP Delivery Date Automation Demo](images/excel-demo.png)

### Demo Files

[View Demo Workbook](demo/SAP_delivery_date_automation_demo.xlsx)

[View VBA Source Code](src/SAP_delivery_date_automation_demo.bas)

---

## Technologies

- SAP ERP / ECC
- SAP GUI Scripting
- Excel VBA
- Order Management
- Process Automation

---

*Portfolio version — company, customer, and production data have been removed or replaced with fictional examples.*
