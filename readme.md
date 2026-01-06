# **Excel Automation & Data Visualization Pipeline**

A Python-based utility designed to automate repetitive spreadsheet tasks, specifically for price adjustments and automated reporting. This script processes large datasets in Excel, performs calculations, and generates visual charts automatically.

## **🚀 Overview**

Manually updating prices and generating charts in Excel is time-consuming and prone to human error. This script leverages the openpyxl library to:

1. **Extract** data from existing workbooks.  
2. **Transform** data by applying a 10% discount across specific product columns.  
3. **Load** results into a new column.  
4. **Visualize** trends by automatically generating and embedding a Bar Chart.

## **🛠️ Tech Stack**

* **Language:** Python 3.x  
* **Library:** openpyxl (Excel integration)

## **📋 Features & Logic**

* **Dynamic Row Handling:** Uses sheet.max\_row to ensure the script works regardless of the dataset size.  
* **Data Validation:** Includes a check for None values to prevent script crashes during calculation.  
* **Automated Charting:** Utilizes the Reference and BarChart classes to define data boundaries and render a chart at position 'f2'.  
* **Non-Destructive Editing:** Saves results to a new file (updated\_file.xlsx), preserving the integrity of the original source data.

## **🖥️ Usage**

1. Place your source Excel file (e.g., Python Automatic.xlsx) in the project directory.

Run the script:  
Bash  
python process\_workbook.py

2.   
3. Open updated\_file.xlsx to see the new prices in Column 5 and the generated Bar Chart.

## **📊 Sample Transformation**

| Original Price (Col 3\) | Processed Price (Col 5\) | Chart Generated |
| :---- | :---- | :---- |
| 100.00 | 90.00 | Yes |
| 50.00 | 45.00 | Yes |

