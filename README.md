# Customer Order Dashboard – Excel Automation Project

## Overview

**Anant Foods Pvt. Ltd.** is an Indian B2B wholesale distributor supplying groceries and food products to retail stores across South India.  
To streamline client relationship management and improve sales follow-up efficiency, the company needed a centralized view of customer order history and engagement patterns.

This **Excel-based Customer Order Dashboard** enables the sales and operations teams to quickly access key customer details such as contact information, location, total number of orders, average freight cost, and last order date.  
By leveraging **Excel formulas, advanced filters, and VBA automation**, the tool provides a real-time, interactive experience for customer insights, all within a single workbook.

---

## Business Questions

- How many orders has each customer placed to date?  
- What is the average freight cost per customer?  
- When was the most recent order placed for each customer?  
- Which shipping methods (UPS, DHL, etc.) are most commonly used?  
- How can the sales team quickly filter and analyze customer order history?  
- Which customers show consistent purchasing activity (high engagement)?  
- How can this data improve follow-ups and customer retention?

---

## Tools and Methodology

### Tools
- **Microsoft Excel:** Dashboard design, formulas, and layout  
- **VBA Macros:** Automation of data filtering and dashboard refresh  
- **Advanced Filters:** For dynamic order history extraction  
- **Data Validation:** For dropdown-based customer selection  
- **Formulas Used:** `VLOOKUP`, `XLOOKUP`, `index and Match`, `PROPER`, `UPPER`, and conditional formatting  

### Methodology

**Data Preparation:**  
Raw sales order data was cleaned and standardized. Customer IDs, order dates, and shipment details were verified for consistency. Text functions (`PROPER`, `UPPER`) were used to normalize formatting across datasets.

**Data Processing:**  
Lookups (`VLOOKUP` / `XLOOKUP`) dynamically retrieve customer information when a name is selected from a dropdown list.  
Key metrics such as *Order Count*, *Average Freight*, and *Last Order Date* are automatically calculated.

**Automation:**  
A **VBA Macro** was written to automate the *Advanced Filter* function. When the user clicks the “Advanced Filter” button, the dashboard instantly updates the Order History table based on the selected customer.

**Visualization:**  
The dashboard summarizes data with neatly formatted sections for:  
- Customer Details  
- Key Metrics (Order Count, Freight, Last Order Date)  
- Filtered Order History Table  

---

## Dashboard Preview  

![Customer Order Dashboard](https://github.com/dorischioma/Customer-Order-Dashboard-Excel-Automation-Project/blob/main/Customer%2BOrders.img.png)

