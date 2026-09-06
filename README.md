# 📊 FUNDAMENTAL BOOSTER | EXCEL PROJECT

> 🚀 A practical Excel project demonstrating fundamental and intermediate Excel skills using Employee, Sales, and Student datasets.

---

## 📌 Project Overview

This project is created as part of **PR. 1 – Fundamental Booster** to practice and demonstrate important Excel concepts used in real-world data analysis.

The main objective of this project is to understand how Excel formulas and functions can be used to:

- Analyze data
- Apply logical conditions
- Perform calculations
- Search and retrieve information
- Clean and transform data
- Work with dates
- Create dynamic results
- Generate meaningful insights from raw datasets

---

## 📁 Project Structure

The Excel workbook contains the following sheets:

| Sheet | Description |
|---|---|
| 📖 Project Guide | Complete project topics and task requirements |
| 👥 Employee Data | Employee information, salary and joining-date analysis |
| 💰 Sales Data | Product, region, salesperson, amount and sales analysis |
| 🎓 Students | Student marks, averages, grades and performance analysis |
| 📝 Task Solutions | Main Tasks and Essential Tasks with formulas and expected results |


---

# 🎯 Main Tasks

## 1️⃣ Student Grade Classification

Student averages are calculated using:

    =AVERAGE(C2:E2)

Grades are assigned using Nested IF:

    =IF(G2>=90,"A",IF(G2>=75,"B",IF(G2>=60,"C","F")))

### Grading Criteria

| Average Score | Grade |
|---:|:---:|
| 90+ | A |
| 75–89 | B |
| 60–74 | C |
| Below 60 | F |

This demonstrates the use of **AVERAGE and Nested IF functions**.

---

## 2️⃣ COUNTIFS – Student Performance

To count students scoring above 60 in Mathematics:

    =COUNTIFS(Students!C2:C21,">60")

### Result

**19 students**

COUNTIFS is useful when we need to count records based on one or more conditions.

---

## 3️⃣ VLOOKUP – Sales Lookup

Example: Find the sales amount for Sales ID `2008`:

    =VLOOKUP(2008,'Sales Data'!A2:F21,5,FALSE)

### Result

**₹46,429**

VLOOKUP searches for a value in the first column of a range and returns related information from another column.

---

## 4️⃣ XLOOKUP – Employee Salary

Example: Find the salary of Employee ID `3018`:

    =XLOOKUP(3018,'Employee Data'!A2:A21,'Employee Data'!D2:D21,"Not Found")

### Result

**₹116,284**

XLOOKUP is a modern and flexible alternative to VLOOKUP.

---

## 5️⃣ Joining Date Analysis

Employee joining dates are analyzed using different Excel date functions.

### Joining Year

    =YEAR(E2)

### Joining Month

    =TEXT(E2,"mmmm")

### Joining Day

    =DAY(E2)

These functions convert date information into useful analytical fields.

---

## 6️⃣ FILTER – Top Performing Students

Students with an average score above 80 can be extracted using:

    =FILTER(Students!A2:H21,Students!G2:G21>80,"No students found")

This demonstrates the use of Excel's **Dynamic Array and FILTER functions**.

---

# ⭐ Essential Tasks

## 1. Relative & Absolute Cell References

Example:

    =D2*(1+$J$2)

Here:

- `D2` → Relative Reference
- `$J$2` → Absolute Reference

When the formula is copied down:

- `D2` changes to `D3`, `D4`, etc.
- `$J$2` remains fixed.

This demonstrates the difference between **Relative and Absolute References**.

---

## 2. IF Formula

Example:

    =IF(G2>=60,"Pass","Fail")

The IF function is used to make decisions based on conditions.

---

## 3. Nested IF

Student grades are classified using:

    =IF(G2>=90,"A",IF(G2>=75,"B",IF(G2>=60,"C","F")))

Nested IF is useful when multiple conditions need to be checked.

---

## 4. IF with AND

To identify students who scored above 80 in both Math and Science:

    =IF(AND(C2>80,D2>80),"Yes","No")

This checks whether **both conditions are TRUE**.

---

## 5. IF with OR

Example:

    =IF(OR(E2>=40000,B2="Laptop"),"Eligible","Not Eligible")

If either condition is TRUE, the result is **Eligible**.

---

## 6. COUNTIFS

To count students who scored above 50 in Mathematics:

    =COUNTIFS(Students!C2:C21,">50")

### Result

**19 students**

---

## 7. SUMIFS

To calculate total sales for the **South region and Keyboard product**:

    =SUMIFS('Sales Data'!E2:E21,'Sales Data'!C2:C21,"South",'Sales Data'!B2:B21,"Keyboard")

### Result

**₹42,233**

SUMIFS is useful for calculating totals based on multiple conditions.

---

## 8. AVERAGEIFS

To calculate the average of student averages above 60:

    =AVERAGEIFS(Students!G2:G21,Students!G2:G21,">60")

### Result

**73.625**

---

## 9. VLOOKUP

### Find Student Name by Student ID

    =VLOOKUP(1008,Students!A2:F21,2,FALSE)

### Result

**Student 8**

---

## 10. INDEX + MATCH

To find the salary of Employee ID `3016`:

    =INDEX('Employee Data'!D2:D21,MATCH(3016,'Employee Data'!A2:A21,0))

### Result

**₹114,173**

### How it works

**MATCH** finds the position of the Employee ID.

**INDEX** returns the salary from that position.

This combination is useful for flexible lookups.

---

## 11. INDEX + MATCH with Multiple Criteria

To find sales for a specific salesperson in a specific month:

    =INDEX('Sales Data'!E2:E21,MATCH(1,('Sales Data'!D2:D21="Person 10")*('Sales Data'!G2:G21="May"),0))

### Example Result

**₹35,185**

This demonstrates lookup using multiple conditions.

---

## 12. TEXT Functions

### Extract First Name

    =LEFT(B2,FIND(" ",B2)-1)

### Convert Text to Uppercase

    =UPPER(B2)

### Convert Text to Lowercase

    =LOWER(B2)

These functions are useful for **data cleaning and standardization**.

---

## 13. XLOOKUP

Example:

    =XLOOKUP(3016,'Employee Data'!A2:A21,'Employee Data'!D2:D21,"Not Found")

### Result

**₹114,173**

XLOOKUP can search data without the column restrictions of traditional VLOOKUP.

---

## 14. XMATCH

To find the position of Laptop in the Sales Data product list:

    =XMATCH("Laptop",'Sales Data'!B2:B21,0)

### Result

**13**

XMATCH returns the position of the matching value inside the selected range.

---

## 15. INDIRECT

Example:

    =INDIRECT("'Employee Data'!"&"D5")

INDIRECT creates a cell reference dynamically from text.

This can be useful when cell or range references need to change dynamically.

---

## 16. OFFSET

Example:

    =SUM(OFFSET('Sales Data'!E2,0,0,5,1))

This creates a dynamic range starting from the selected cell.

OFFSET can be useful for dynamic calculations and analysis.

---

## 17. Date & Time Functions

### YEAR

    =YEAR(E2)

### MONTH

    =MONTH(E2)

### DAY

    =DAY(E2)

### Format Date

    =TEXT(E2,"dd-mmm-yyyy")

Example:

    01-Jan-2015

---

## 18. Date Difference

To calculate the number of days between two dates:

    =E21-E2

### Example Result

**4750 days**

This can be useful for calculating durations between dates.

---

## 19. Age Calculation

If a Date of Birth column is available:

    =DATEDIF(DOB,TODAY(),"Y")

This calculates the person's age in completed years.

---

## 20. Mathematical Functions

### ROUND

    =ROUND(E2,-2)

Example:

    7048 → 7000

Rounds a number to the nearest hundred.

---

### CEILING

    =CEILING(E2,1000)

Example:

    7048 → 8000

Rounds a number upward to the nearest thousand.

---

### FLOOR

    =FLOOR(E2,1000)

Example:

    7048 → 7000

Rounds a number downward to the nearest thousand.

---

# 💰 Business Logic – Discount Calculation

Sales discounts are calculated according to the sales amount.

| Sales Amount | Discount |
|---:|---:|
| ₹40,000+ | 20% |
| ₹20,000–₹39,999 | 10% |
| Below ₹20,000 | 5% |

Formula:

    =IF(E2>=40000,20%,IF(E2>=20000,10%,5%))

Discount Amount:

    =E2*H2

This demonstrates how Excel can be used to implement **real-world business rules**.

---

# 📊 Sample Project Results

| Analysis | Result |
|---|---:|
| Students with Math > 60 | **19** |
| Students with Math > 50 | **19** |
| South + Keyboard Total Sales | **₹42,233** |
| Employee 3018 Salary | **₹116,284** |
| Employee 3016 Salary | **₹114,173** |
| Sales ID 2008 Amount | **₹46,429** |
| Person 10 – May Sales | **₹35,185** |

---

# 🛠️ Excel Functions Used

    AVERAGE
    IF
    AND
    OR
    COUNTIFS
    SUMIFS
    AVERAGEIFS
    VLOOKUP
    XLOOKUP
    INDEX
    MATCH
    XMATCH
    LEFT
    FIND
    UPPER
    LOWER
    INDIRECT
    OFFSET
    YEAR
    MONTH
    DAY
    TEXT
    DATEDIF
    ROUND
    CEILING
    FLOOR
    FILTER

---

# 🧠 What I Learned

Through this project, I practiced:

- ✅ Working with structured datasets
- ✅ Writing Excel formulas
- ✅ Applying logical conditions
- ✅ Performing conditional calculations
- ✅ Searching and retrieving data
- ✅ Using lookup functions
- ✅ Cleaning and transforming text
- ✅ Working with dates
- ✅ Performing mathematical calculations
- ✅ Creating dynamic ranges
- ✅ Using Dynamic Array functions
- ✅ Understanding Relative and Absolute References
- ✅ Applying Excel concepts to real-world business scenarios

---

# ⚠️ Dataset Notes

### Product Price / Product Code

The supplied Sales dataset contains:

- Sales ID
- Product
- Region
- Salesperson
- Amount
- Date

A separate **Product Price/Product Code** table was not provided.

Therefore, the workbook demonstrates the lookup concept using the available Sales ID and Amount data.

---

### Date of Birth

The supplied Employee and Student datasets do not contain a Date of Birth column.

Therefore, the age calculation is documented using:

    =DATEDIF(DOB,TODAY(),"Y")

This can be applied once a DOB column is available.

---

# 📂 Project File

The completed Excel workbook contains:

    PR_1_Fundamental_Booster_Completed_Final.xlsx

The workbook includes:

- 📖 Project Guide
- 👥 Employee Data
- 💰 Sales Data
- 🎓 Students
- 📝 Task Solutions
- 📊 Completed formulas
- 📌 Expected results
- 🎨 Project formatting

---

# 🎯 Project Objective

> **The objective of this project is to demonstrate practical Excel skills by applying formulas and functions to Employee, Sales, and Student datasets for data analysis, lookup, transformation, and decision-making.**

---

# 🚀 Future Improvements

This project can be further enhanced by adding:

- 📊 Pivot Tables
- 📈 Interactive Excel Dashboard
- 🎯 KPI Cards
- 📉 Sales Charts
- 🔍 Slicers
- 📋 Data Validation
- 🧹 Advanced Data Cleaning
- 📊 Conditional Formatting
- 📈 Advanced Business Analysis

---

# 👨‍💻 Author

## Rakesh

**Skills:**  
Excel | SQL | Power BI | Python | Data Analysis

---

⭐ If you found this project useful, feel free to give it a star!

> **"Learning Excel is not just about formulas — it's about turning raw data into meaningful insights."**
