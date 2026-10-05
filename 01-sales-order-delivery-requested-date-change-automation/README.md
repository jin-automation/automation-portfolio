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
