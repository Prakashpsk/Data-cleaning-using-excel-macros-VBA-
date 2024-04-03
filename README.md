# Data Cleaning Using Excel Macros (VBA)

## INTRODUCTION
This project aims to streamline regular tasks in Excel by automating them using macros. In this example, we'll walk through a data cleaning task step by step.

## Overview
1. **Download the Data from a Webpage**: The downloaded file is a text file with comma-separated values.
   ![Text Data](https://github.com/Prakashpsk/Data-cleaning-using-excel-macros-VBA-/blob/main/text_data.png)

2. **Convert Text File to CSV**: If needed, convert the text file to a CSV file.
   ![Raw Data](https://github.com/Prakashpsk/Data-cleaning-using-excel-macros-VBA-/blob/main/raw_data.png)

3. **Use Macros (VBA) to Automate Daily Tasks**: We'll utilize VBA macros to automate the data cleaning process.

## Procedure
1. **Download the Data from a Webpage**: Obtain the data from the webpage as a text file.
2. **Convert Text File to CSV**: If necessary, convert the text file to a CSV file for easier handling in Excel.
3. **Use Macros (VBA) Code**: Implement VBA code to automate the data cleaning process. Below is a basic example:

   ```vba
   Sub SplitDataByComma()
       Dim rng As Range
       Dim cell As Range
       Dim splitData() As String
       Dim i As Integer

       ' Define the range containing the data
       Set rng = Range("A1:A" & Cells(Rows.Count, "A").End(xlUp).Row)

       ' Loop through each cell in the range
       For Each cell In rng
           ' Split the cell value by comma separator
           splitData = Split(cell.Value, ",")

           ' Copy the split data to adjacent cells
           For i = LBound(splitData) To UBound(splitData)
               cell.Offset(0, i).Value = splitData(i)
           Next i
       Next cell
   End Sub
This macro splits the data in column A into multiple columns based on the comma separator.

Final Output
To use this macro:

Press Alt + F11 to open the Visual Basic for Applications (VBA) editor.
Insert a new module from the "Insert" menu.
Copy and paste the above VBA code into the module.
Close the VBA editor.
Press Alt + F8, select "SplitDataByComma", and click "Run".
This macro will split the data in column A into multiple columns based on the comma separator.

final outpout
![](https://github.com/Prakashpsk/Data-cleaning-using-excel-macros-VBA-/blob/main/aftter_macrows_data.png)

Choose the method that best suits your needs and level of automation required.
