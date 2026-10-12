# Order Tracking & Delivery Status Automation

**Excel VBA | Order Management | Logistics Automation**

Automated order and delivery tracking by matching Customs Clearance List and POD (Proof Of Delivery) data against internal order lists and updating order status automatically in Excel.

---

## Business Problem

Order status required manual comparison between internal order lists, Customs Clearance lists, and POD data.

For large order volumes, this was repetitive and time-consuming.

---

## Solution

I developed Excel VBA automations for two key processes:

### Customs Clearance List Matching
- Matches Sales Orders against the Customs Clearance List.
- Identifies matching orders automatically.
- Updates relevant order information and status.

### POD Delivery Confirmation
- Matches POD data against internal Order Numbers.
- Updates completed orders with **Status and POD Date** automatically.

### Data Flow

**Customs Clearance / POD Data → Order Matching → Status Update**

---

## Business Impact

- Reduced repetitive manual file comparison.
- Automated order and delivery status updates.
- Reduced manual Excel entry.

---

## Demo

![Order Tracking Automation Demo](./images/excel-demo.png)

[View Demo Workbook](./demo/Order_tracking_automation_demo.xlsx)

[View Customs Clearance Matching Code](./src/Customs_clearance_matching.bas)

[View POD Delivery Confirmation Code](./src/POD_delivery_confirmation.bas)

---

*Portfolio version — company, customer, order, customs clearance, and delivery data have been replaced with fictional examples.*
