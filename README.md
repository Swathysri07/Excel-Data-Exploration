# Data Exploration and Analysis

## 1. Project Overview

This project demonstrates basic data exploration and analysis using Microsoft Excel. The analysis is performed on a Product Dataset containing product details such as Product ID, Product Name, Brand Name, Quantity, Category, and Price.

The project focuses on numerical calculations, statistical analysis, logical conditions, conditional calculations, and text extraction using built-in Excel functions.

## 2. Dataset Description

**Dataset Name:** Product Dataset

The dataset contains product-level information. Each row represents a product record, and each column represents a specific product attribute.

### Dataset Attributes

| Column Name | Description |
|---|---|
| Product ID | Unique identification code assigned to each product. |
| Product Name | Name or description of the product. |
| Brand Name | Brand associated with the product. |
| Quantity | Quantity value recorded for the product. |
| Category | Classification of the product, such as Electronics. |
| Price | Price of the product in dollars ($). |

### Purpose of the Dataset

The dataset provides product information for analysing prices, counting product records, comparing price values, identifying products by category, and extracting information from Product IDs.

## 3. Project Objectives

- Calculate total, count, and average product prices.
- Identify minimum and maximum product prices.
- Categorize products based on price conditions.
- Calculate the total price of products in a specific category.
- Count products that meet a specified price condition.
- Extract Day, Month, and Country Code information from Product IDs.

## 4. Excel Functions and Formula Explanation

The following Excel functions are applied to perform the analysis.

### A. Basic Data Exploration

**1. SUM Function**

Formula:
`=SUM(price_range)`

Purpose: Calculates the total price of all products by adding the values in the Price column.

**2. COUNT Function**

Formula:
`=COUNT(price_range)`

Purpose: Counts the number of numerical entries in the Price column, representing the number of products with numeric price values.

**3. AVERAGE Function**

Formula:
`=AVERAGE(price_range)`

Purpose: Calculates the average price of all products by dividing the sum of prices by the number of numeric price entries.

### B. Minimum and Maximum Analysis

**4. MIN Function**

Formula:
`=MIN(price_range)`

Purpose: Identifies the lowest price available in the dataset.

**5. MAX Function**

Formula:
`=MAX(price_range)`

Purpose: Identifies the highest price available in the dataset.

### C. Logical Function – IF

**6. IF Function**

Formula:
`=IF(price_cell>=500,"High Price","Standard Price")`

Purpose: Creates a new column named **Price Range** and classifies products based on their price.

Conditions:
- Price greater than or equal to $500 → High Price
- Price less than $500 → Standard Price

### D. Conditional Functions

**7. SUMIF Function**

Formula:
`=SUMIF(category_range,"Electronics",price_range)`

Purpose: Calculates the total price of products belonging to the Electronics category.

**8. COUNTIF Function**

Formula:
`=COUNTIF(price_range,"<100")`

Purpose: Counts the number of products with a price below $100.

### E. Text Extraction Functions

These functions extract specific characters from the Product ID column and create additional columns.

**9. LEFT Function**

Formula:
`=LEFT(product_id_cell,2)`

Purpose: Extracts the first two characters from the Product ID and stores the result in a new column named **Day**.

**10. RIGHT Function**

Formula:
`=RIGHT(product_id_cell,2)`

Purpose: Extracts the last two characters from the Product ID and stores the result in a new column named **Country Code**.

**11. MID Function**

Formula:
`=MID(product_id_cell,4,3)`

Purpose: Extracts three characters starting from the fourth position of the Product ID and stores the result in a new column named **Month**.

## 5. Data Transformation

The following additional columns are created using Excel text and logical functions:

| New Column | Function Used | Transformation |
|---|---|---|
| Price Range | IF | Categorizes products into High Price and Standard Price. |
| Day | LEFT | Extracts the first two characters of Product ID. |
| Country Code | RIGHT | Extracts the last two characters of Product ID. |
| Month | MID | Extracts the 4th, 5th, and 6th characters of Product ID. |

## 6. Workbook Structure

The Excel workbook is organized into three worksheets:

- **RAW DATA:** Contains the original product dataset.
- **DATA EXPLORATION:** Contains the dataset with additional columns created through Excel functions.
- **CALCULATION SHEET:** Contains the calculated results using numerical, statistical, logical, and conditional functions.

## 7. Tools and Technologies

- Microsoft Excel
- Excel Formulas and Built-in Functions

## 8. Skills Demonstrated

- Data Exploration
- Data Summarization
- Statistical Analysis
- Logical Analysis
- Conditional Aggregation
- Text Manipulation
- Data Transformation
- Spreadsheet Management

## 9. Conclusion

This project covers fundamental data exploration and analysis techniques using Microsoft Excel. It includes numerical calculations, statistical functions, price-based categorization, conditional calculations, and text extraction from product identifiers.

The completed workbook demonstrates the application of Excel functions to organize, transform, and analyse product data.
