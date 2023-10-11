# NHans_UCI

The code and screenshots are included in this repository.

----

I found a numerous examples of worksheet looping online that was helpful in looping through the full workbook.

—

Dim ws As Worksheet

For Each ws In Sheets MsgBox ws.Name

—

Similarly, I reused the lastrow function from the census exercise, modified to fit the purpose of this exercise

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
