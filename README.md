Excel Data Exploration – Product Dataset

A beginner-level data analysis project in *Microsoft Excel* that explores a product dataset using core formulas: summary statistics, logical functions, conditional aggregation, and text extraction.



## 📌 Project Overview

As a data analyst, the goal is to explore a product dataset and prepare it for further analysis by:

- Summarizing numeric data (total, count, average, min, max)
- Categorizing products with business logic (IF)
- Aggregating data by condition (SUMIF, COUNTIF)
- Extracting structured information from text (LEFT, RIGHT, MID)

## 🗂️ Dataset

The dataset contains *34 product records* with the following columns:

| Column | Description |
|---|---|
| Product ID | Coded ID in the format DD-MON-CC (day, month, country code), e.g. 28-JAN-US |
| Product Name | Name of the product (Laptop, Sneakers, Blender, etc.) |
| Brand Name | Brand of the product (Dell, Nike, Sony, etc.) |
| Price ($) | Price of the product in US dollars |
| Quantity | Quantity of the product |
| Category | Electronics, Fashion, Kitchen, Outdoor, Accessories |

*Columns added during analysis:* Price Range, Day, Country Code, Month

## 🛠️ Skills & Excel Functions Used

| Skill | Functions |
|---|---|
| Data summarization | SUM, COUNT, AVERAGE |
| Range analysis | MIN, MAX |
| Logical analysis | IF |
| Conditional aggregation | SUMIF, COUNTIF |
| Text manipulation | LEFT, RIGHT, MID |

---

## 🔍 Tasks, Formulas & Results

### 1. Sum, Count and Average

| Metric | Formula | Result |
|---|---|---|
| Total price of all products | =SUM(D2:D35) | *10,100* |
| Number of products | =COUNT(D2:D35) | *34* |
| Average price | =AVERAGE(D2:D35) | *≈ 297.06* |

### 2. Minimum and Maximum Price

| Metric | Formula | Result |
|---|---|---|
| Minimum price | =MIN(D2:D35) | *30* |
| Maximum price | =MAX(D2:D35) | *1,000* |

### 3. Logical Function – IF (Price Range column)

Products are classified as *High Price* if Price ≥ $500, otherwise *Standard Price*.

excel
=IF(D2>=500,"High Price","Standard Price")


Result: *8 High Price* products and *26 Standard Price* products.

### 4. Conditional Functions – SUMIF and COUNTIF

| Task | Formula | Result |
|---|---|---|
| Total price of Electronics products | =SUMIF(F2:F35,"Electronics",D2:D35) | *8,050* |
| Number of products priced below $100 | =COUNTIF(D2:D35,"<100") | *11* |

### 5. Text Functions – LEFT, RIGHT, MID

Information extracted from the Product ID column (example: 28-JAN-US):

| New Column | Formula | Example Output |
|---|---|---|
| Day | =LEFT(A2,2) | 28 |
| Country Code | =RIGHT(A2,2) | US |
| Month | =MID(A2,4,3) | JAN |

---

## 📁 Repository Structure


excel-data-exploration/
│
├── Excel_Assignment_1_-_Data_Exploration.xlsx   # Workbook with formulas and new columns
├── Excel_Assignment_1_Documentation.pdf         # Screenshots of results and formula bar
├── screenshots/                                 # (optional) individual screenshots
└── README.md


## 📸 Screenshots

> Add your screenshots to the screenshots/ folder and link them here, for example:

markdown
![Summary statistics](screenshots/summary_statistics.png)
![Price Range column](screenshots/price_range.png)
![Text functions](screenshots/text_functions.png)


## 💡 Key Insights

- The average price (~$297) is well above the median price ($140), since a handful of high-priced items (laptops, smartphones, camera, tablet) pull the mean up.
- *Electronics* is the dominant category by value: $8,050 of the $10,100 total (~80%).
- About *one third* of products (11 of 34) are priced below $100.
- Product IDs embed# Excel-Assignment--1
