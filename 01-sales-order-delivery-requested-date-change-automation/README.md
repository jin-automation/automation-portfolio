# Sales Order Delivery Requested Date Change Automation

**SAP ERP/ECC | Excel VBA | SAP GUI Scripting | Order Management | Process Automation**

> Automated high-volume sales order delivery date changes in SAP ECC, reducing a 40+ minute repetitive manual task to approximately 10 minutes while adding validation controls to reduce human error.

---

## Overview

This automation was developed to solve a repetitive and time-consuming process in Sales Order Management.

One of our key resellers frequently requested delivery date changes for a large number of sales order line items. A single request could involve more than **100 line items**.

The reseller often provided only their own **Purchase Order (PO) numbers**, rather than our internal Sales Order numbers.

As a result, each line item had to be processed manually by searching SAP using the customer's PO number, locating the corresponding Sales Order and line item, verifying the requested item, and changing the delivery requested date.

For requests involving more than 100 line items, this repetitive process could take **more than 40 minutes** to complete.

---

## Business Problem

The original manual process required the Order Management user to repeatedly:

1. Open the Sales Order change transaction in SAP.
2. Search for the Sales Order using the customer's PO number.
3. Locate the requested line item.
4. Verify the Material / Part Number.
5. Verify the Order Quantity.
6. Change the requested delivery date.
7. Save the Sales Order.
8. Repeat the same process for the next line item.

For a large reseller request containing 100+ line items, these steps had to be repeated more than 100 times.

Besides being time-consuming, the repetitive nature of the process increased the risk of human error, including:

- Selecting the wrong Sales Order
- Selecting the wrong line item
- Updating the wrong Material / Part Number
- Updating an item with a different quantity
- Entering an incorrect delivery date
- Missing one or more line items during repetitive processing

---

## Solution

I developed an **Excel VBA automation integrated with SAP GUI Scripting** to automate the complete process.

Instead of manually processing each request in SAP, the required information could be prepared in Excel and the automation started with a single action.

The automation then processes each line item sequentially.

For every request, it:

1. Reads the customer's PO number from Excel.
2. Opens the SAP Sales Order change process.
3. Searches for the corresponding Sales Order using the customer's PO number.
4. Navigates to the requested line item.
5. Retrieves the Material / Part Number from SAP.
6. Retrieves the Order Quantity from SAP.
7. Compares the SAP Material Number against the expected Material Number in Excel.
8. Compares the SAP Order Quantity against the expected quantity in Excel.
9. Updates the requested delivery date **only when both values match**.
10. Saves the Sales Order change.
11. Records the processing result in Excel.
12. Continues automatically to the next line item.

---

## Automation Workflow

```text
Excel Request Data
        |
        v
Customer PO + Line Item + Material + Qty + New Date
        |
        v
      SAP ECC
        |
        v
Search Sales Order using Customer PO
        |
        v
Locate Requested Line Item
        |
        v
Read SAP Material Number & Quantity
        |
        v
Compare SAP Data with Excel Request
        |
        v
   Material & Qty Match?
       /          \
     NO            YES
      |              |
      v              v
Skip Update      Change Requested
& Log Reason      Delivery Date
      |              |
      |              v
      |           Save in SAP
      |              |
      \______________/
             |
             v
      Record Status in Excel
             |
             v
      Process Next Line Item
```

---

## Safety & Validation

One of the most important considerations when designing this automation was preventing incorrect updates in SAP.

The automation therefore does **not** simply locate a line item and change its date.

Before making any change, it verifies two important values:

- **Material / Part Number**
- **Order Quantity**

The delivery requested date is changed only when both SAP values match the expected values provided in Excel.

If either value does not match, the automation skips the update and records the reason for review.

Examples of recorded statuses include:

```text
OK
MISMATCH: PN
MISMATCH: Qty
MISMATCH: PN & Qty
SKIP: Missing data
NO DATE
DATE WRITE FAIL
ERROR
```

This validation mechanism was designed to reduce the risk of modifying the wrong Sales Order line item during high-volume processing.

---

## Business Impact

### Before Automation

- Large requests could contain **100+ line items**.
- Processing 100+ line items manually could take **more than 40 minutes**.
- Each item required repetitive SAP navigation, verification and data entry.
- The reseller's PO number often had to be used to locate the corresponding internal Sales Order.
- Repetitive manual processing increased the possibility of human error.
- Large delivery-date-change requests required significant attention and manual effort.

### After Automation

- The same type of high-volume request could typically be processed in **approximately 10 minutes**.
- More than 100 line items could be processed automatically after starting the automation.
- Minimal user intervention was required while the automation was running.
- Material Number and Quantity validation helped prevent incorrect SAP updates.
- Exceptions and mismatches were automatically recorded for manual review.
- Large-volume reseller requests became significantly easier and faster to manage.

---

## Before vs After

| Process | Before Automation | After Automation |
|---|---|---|
| 100+ line-item request | 40+ minutes | ~10 minutes |
| Sales Order search | Manual | Automated using Customer PO |
| Line-item navigation | Manual | Automated |
| Material verification | Manual | Automated |
| Quantity verification | Manual | Automated |
| Delivery date update | Manual | Automated after validation |
| Error checking | User dependent | Built-in validation |
| Exception tracking | Manual | Automatically logged in Excel |
| User intervention | Continuous | Minimal after execution |

---

## Excel Input & Processing Log

The Excel worksheet acts as both the input source and processing log.

Typical fields include:

| Field | Purpose |
|---|---|
| Customer PO | Used to locate the Sales Order |
| Line Item | Identifies the requested Sales Order line |
| Material / Part Number | Used for validation |
| Quantity | Used for validation |
| Requested Delivery Date | New date to be updated |
| PN / Qty Match | Indicates validation result |
| Status | Records processing result |
| SAP Material | Captures SAP value for verification |
| SAP Quantity | Captures SAP value for verification |

This provides visibility into which items were successfully processed and which items require manual review.

---

## Technologies Used

- **SAP ERP / ECC**
- **SAP VA02**
- **SAP GUI Scripting**
- **Microsoft Excel**
- **Excel VBA / Macros**
- **Business Process Automation**
- **Order Management**

---

## Key VBA Concepts Used

The automation demonstrates practical use of:

- SAP GUI Scripting through VBA
- Excel-to-SAP integration
- Automated SAP transaction navigation
- Dynamic processing of multiple Excel rows
- Material Number normalization
- Quantity normalization and comparison
- Conditional SAP updates
- Error handling
- Exception logging
- Automated status reporting

---

## Key Design Principle

The objective was not simply to make the process faster.

The automation was designed around the principle:

> **Validate first, update second.**

A Sales Order line is updated only after the automation confirms that the Material Number and Order Quantity in SAP match the expected request.

This combines **process efficiency with operational control**, which is particularly important when automating changes to live ERP transactions.

---

## Result

This project transformed a repetitive Sales Order maintenance task from a manual, high-attention process into a largely automated workflow.

**100+ Line Items | 40+ Minutes → ~10 Minutes | SAP ECC + Excel VBA | Built-in Validation**

The automation allowed high-volume delivery date change requests to be handled more efficiently while reducing repetitive manual work and the risk of incorrect updates.

---

## Source Code

A sanitized portfolio version of the VBA source code will be included in this repository.

All company-specific, customer-specific and confidential information will be removed or generalized before publication.

---

## Note

This repository is intended to demonstrate the automation approach, business problem-solving process and technical implementation.

No confidential company, customer, pricing, order or production SAP data is included.
