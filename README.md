# 📊 Excel Data Cleaning & Transformation – Product Dataset

## 👋 About This Project

This project is part of my **Data Analytics learning journey**, where I am building practical skills in **Microsoft Excel, Data Cleaning, Data Preparation, and Data Transformation**.

As an aspiring **Data Analyst**, I worked with a Product Dataset containing product IDs, manufacturing information, product names, brands, quantities, categories, and prices.

The main objective of this assignment was to transform a raw dataset into a **clean, standardized, structured, and analysis-ready dataset** using Excel.

---

## 🎯 Project Objective

The objective of this project is to demonstrate fundamental data preprocessing techniques required before performing data analysis.

The assignment focuses on:

* Handling missing values
* Identifying inconsistent data
* Standardizing text fields
* Correcting category inconsistencies
* Removing duplicate records
* Splitting columns
* Merging columns
* Formatting numerical and date fields
* Applying conditional formatting
* Preparing the dataset for further analysis

---

## 📁 Dataset Description

The dataset contains product-related information.

### Original Dataset Attributes

| Column       | Description                                                       |
| ------------ | ----------------------------------------------------------------- |
| Product ID   | Unique product identifier containing date and country information |
| Product Name | Name of the product                                               |
| Brand Name   | Brand associated with the product                                 |
| Quantity     | Available quantity                                                |
| Category     | Product category                                                  |
| Price        | Product price                                                     |

### Transformed Columns

| Column             | Purpose                         |
| ------------------ | ------------------------------- |
| Product ID         | Original identifier             |
| Manufacturing Date | Extracted manufacturing date    |
| Country Code       | Extracted country code          |
| Product Name       | Standardized product name       |
| Brand Name         | Brand information               |
| Product Brand      | Combined brand and product name |
| Price ($)          | Price formatted as currency     |
| Quantity           | Product quantity                |
| Category           | Standardized product category   |

---

# 🧹 Data Cleaning Process

## 1. Handling Missing Values

The dataset was checked for missing or incomplete values, particularly in the **Price** and **Category** columns.

### Price

Missing or invalid price values should not be replaced blindly because price is an important business variable.

The handling strategy used/considered was:

* Identify blank or invalid price values.
* Check whether a reliable source exists for the missing value.
* If a valid business rule is available, use an appropriate imputation method.
* Otherwise, retain the value as missing and flag it for further investigation rather than creating an inaccurate price.

### Category

For missing or unknown categories:

* Check the Product Name and Brand Name for clues.
* Use an existing category mapping where possible.
* If the category cannot be determined reliably, label it as **Unknown** rather than making an unsupported assumption.

This helps maintain data quality and transparency.

---

## 2. Correcting Inconsistent Data

### Product Name Standardization

The **Product Name** column was reviewed for inconsistent text formatting.

Excel tools such as:

* Find & Replace
* Text functions
* Consistent capitalization
* Removal of unnecessary spaces

can be used to standardize product names.

Examples of standardized product names include:

* Laptop
* Smartphone
* Headphones
* Coffee Maker
* Fitness Tracker
* Hiking Boots

### Category Standardization

The **Category** column was checked for inconsistent spellings, typos, and unexpected values.

The categories were standardized into consistent values such as:

* Electronics
* Fashion
* Kitchen
* Outdoor
* Accessories
* Unknown

Excel's **Find & Replace** feature was used where corrections were required.

---

# 🔁 3. Removing Duplicate Records

Duplicate records can cause incorrect calculations and misleading analysis.

The dataset was checked using:

**Excel → Data → Remove Duplicates**

The comparison was based on the complete row.

Duplicate records were removed so that each product record appears only once where appropriate.

For example, repeated product records such as identical product, brand, quantity, price, and category combinations were checked during the duplicate-removal process.

---

# 🔀 4. Splitting and Merging Data

## Splitting Product ID

The Product ID contains information in the following structure:

```text
28-JAN-US
```

The Product ID was separated into:

```text
Manufacturing Date → 28-JAN
Country Code       → US
```

Excel text functions such as:

```excel
=LEFT(A2,6)
```

and

```excel
=RIGHT(A2,2)
```

were used to extract the required components.

The unnecessary characters/separators were removed or handled as part of the transformation.

---

## Merging Brand Name and Product Name

The **Brand Name** and **Product Name** fields were combined into a new column called:

### Product Brand

Example:

```text
Dell + Laptop
```

becomes:

```text
Laptop-Dell
```

An Excel concatenation formula was used:

```excel
=CONCAT(D2,"-",E2)
```

This creates a single descriptive field that can be useful for reporting and analysis.

---

# 💰 5. Number and Date Formatting

## Price Formatting

The **Price** column was formatted as currency to improve readability.

Example:

```text
1000
```

is displayed as:

```text
$1,000.00
```

This makes financial values easier to understand and present.

## Manufacturing Date Formatting

The Manufacturing Date was formatted using:

```text
DD-MM-YYYY
```

Example:

```text
28-JAN-2026
```

can be displayed as:

```text
28-01-2026
```

This provides a consistent date format throughout the dataset.

---

# 🎨 6. Conditional Formatting

Conditional formatting was applied to make important patterns easier to identify visually.

## Price Column

A **Data Bar / Color Scale** was applied to the Price column.

This allows high and low prices to be identified quickly without manually comparing every value.

## Category Column

A custom conditional formatting rule was applied to highlight:

```text
Electronics
```

This makes Electronics products visually identifiable within the dataset.

---

# 🛠️ Excel Skills Demonstrated

This project demonstrates the following Data Analytics skills:

### Data Cleaning

* Missing value identification
* Missing value handling strategies
* Data validation
* Error identification

### Data Standardization

* Text standardization
* Category correction
* Find & Replace
* Consistent formatting

### Data Quality Management

* Duplicate detection
* Duplicate removal
* Dataset validation

### Data Transformation

* Splitting columns
* Extracting text
* Merging columns
* Creating calculated fields

### Data Formatting

* Currency formatting
* Date formatting
* Number formatting

### Data Visualization

* Data bars
* Color scales
* Conditional formatting

---

# 📂 Project Structure

```text
Excel-Assignment-2-Data-Cleaning/
│
├── 📊 Data Cleaning and Transformation Assig2.xlsx
│
├── 📄 Data_Cleaning_Assignment_2.pdf
│
└── 📖 README.md
```

### Files Included

**Excel File**

`Data Cleaning and Transformation Assig2.xlsx`

Contains the dataset, cleaning operations, transformations, formulas, formatting, and conditional formatting.

**PDF Documentation**

`Data_Cleaning_Assignment_2.pdf`

Contains screenshots and explanations demonstrating:

* Missing value handling
* Find & Replace
* Duplicate removal
* Column splitting
* Column merging
* Number/date formatting
* Conditional formatting

---

# 📸 Screenshots / Documentation

The PDF documentation demonstrates the major Excel operations performed during this project.

Screenshots include:

1. Missing Value Handling
2. Find & Replace Corrections
3. Duplicate Removal
4. Product ID Splitting
5. Brand + Product Name Merging
6. Price Currency Formatting
7. Manufacturing Date Formatting
8. Price Conditional Formatting
9. Electronics Category Highlighting

---

# 📈 Key Learning Outcomes

Through this assignment, I learned how to:

* Prepare raw data before analysis
* Identify and handle missing information
* Standardize inconsistent datasets
* Remove duplicate records
* Extract information from structured text
* Combine multiple columns into a meaningful field
* Apply appropriate number and date formats
* Use Excel conditional formatting to highlight patterns
* Document data-cleaning steps clearly

These are foundational skills for working with real-world datasets as a **Data Analyst**.

---

# 🧰 Tools Used

* **Microsoft Excel**
* Excel Formulas
* Find & Replace
* Remove Duplicates
* Text Functions
* Conditional Formatting
* Number Formatting
* Date Formatting

---

# 🚀 Future Improvements

As I continue developing my Data Analytics skills, I plan to extend this project by:

* Creating Excel Pivot Tables
* Building interactive dashboards
* Performing exploratory data analysis
* Calculating KPIs
* Creating charts and visualizations
* Applying statistical analysis
* Recreating the analysis using SQL
* Performing the same analysis using Python and Pandas
* Building a complete end-to-end Data Analytics project

---

# 📚 Skills & Concepts

```text
Excel
│
├── Data Cleaning
│   ├── Missing Values
│   ├── Duplicate Removal
│   └── Data Validation
│
├── Data Preparation
│   ├── Text Standardization
│   └── Category Standardization
│
├── Data Transformation
│   ├── Split Columns
│   ├── Merge Columns
│   └── Text Extraction
│
├── Data Formatting
│   ├── Currency
│   └── Date Formatting
│
└── Data Visualization
    └── Conditional Formatting
```

---

# 👨‍💻 About Me

I am an **aspiring Data Analyst** currently developing my skills in:

* Microsoft Excel
* SQL
* Python
* Data Cleaning
* Data Analysis
* Data Visualization
* Power BI

This repository is part of my growing **Data Analytics Portfolio**, where I document my learning projects and practical assignments.

---

## ⭐ Portfolio Journey

This is one of the projects in my Data Analytics learning journey.

More projects will be added as I continue building practical experience in:

**Excel → SQL → Python → Power BI → Data Analytics**

---

## 📌 Project Status

**Completed – Excel Data Cleaning & Transformation**

> This project demonstrates beginner-level data cleaning, preparation, transformation, formatting, and visualization skills using Microsoft Excel.
# Data-cleaning-and-Transformation-Assignment2
