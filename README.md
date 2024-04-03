# Data-cleaning-using-excel-macros-VBA-

INTRODUCTION 
the download the file into webpage the file was one text file but cumma sepreated 
how to make that file make seprate the column voise data or make a data frame 
how to automate the daily task is using macrose 

Procejore
1. download the data in to webpage
  
3. if in case text file convert text file convert CSV File
4. uis macrose VBA Code This code will be your daily task will be automate
   If you want to automate this process further or if you need to perform the task frequently, you can use a VBA macro. Here's a basic example:

   
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


To use this macro:

Press Alt + F11 to open the Visual Basic for Applications (VBA) editor.
Insert a new module from the "Insert" menu.
Copy and paste the above VBA code into the module.
Close the VBA editor.
Press Alt + F8, select "SplitDataByComma", and click "Run".
This macro will split the data in column A into multiple columns based on the comma separator.

Choose the method that best suits your needs and level of automation required.
