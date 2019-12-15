# Testing Python Programs with Excel

This was an exercise to practice using Python with Excel. It was prompted by the [task](#acknowledgments):

**"Given a file directory, find all Excel workbooks. Return a new csv file with each row being a sum of: all values in that row, for all sheets, for all workbooks."**

My goal was to practice using Python; explore the Excel-related capabilities of the python data analysis library, pandas; and learn more about how I could use python programming for data analysis.


## Contents

The repository represents the file directory. The directory contains:

* 5 Excel workbooks analyzed in the program 
* *excel_wb.py* - the Python program for finding the Excel sums
* *excel_wb_v2.py* - the program version 2 that implements more pandas internal functions
* *row_sums.csv* - the csv file created by the program

## Progress in the Works

* **Resolve *sum_sheet_rows(df)* discrepancy**: 
Further investigation/testing is needed to see why the pandas sum function used in excel_wb_v2.py returns different results from the custom sum function used in excel_wb.py.
(This difference occurs in the definition of sum_sheet_rows(df).)

* **Clean up with functional programming**:
It would be helpful to rework some of the functions into "pure functions" so that it's clear what they are taking in and what they are returning. This could also mean removing the global variable so the program can be read unambiguously from the main function.
I recently was introduced (or maybe reintroduced?) to functional programming from this video: ["Learning Functional Programming with JavaScript" by Anjana Sofia Vakil (github.com/vakila/)](https://www.youtube.com/watch?v=e-5obm1G_FY)


* **Add further documentation. Clean up script comments and output text.**

* **Add the ability for users to enter the desired file directory into the console**:
Currently, the program is hardcoded to search the file directory containing the program itself.

## Running Requirements

Python 3 installation is needed to run the python file locally.

To run the program, simply download *excel_wb.py* or *excel_wb_v2.py* and save it to a folder containing Excel workbooks.

Run the program through a terminal application (like Mac's Terminal). 
In Terminal, navigate to the folder where the program is located. Then run the file.

```
python3 excel_wb.py
```

## Acknowledgments

* Shout-out to Rohit for giving me "homework" and sending me educational YouTube videos. 
