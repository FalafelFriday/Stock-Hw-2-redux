# Stock-Hw-2-redux

Prerequisites: Microsoft excel with Macro writting enabled 


Coding step 1: The first part was writting this line of code (lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row).This allows us to effectively parse 
through the whole document automatically by row. While defining the variables I had to include a placeholder variable(called P_P) it is used
to make sure i can track the yearly open. Next I set the for loop, and wrote a formula to calculate total volume,percent change total yearly change and stock name
SAMPLE:
---------------------------------------------------------
For i = 2 To lastrow
open_yearly = ws.Cells(P_P, 3).Value
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Stock_name = ws.Cells(i, 1).Value
total_yearly = total_yearly + (Cells(i, 6).Value - open_yearly)
If open_yearly = 0 Then
percent_change = 0
Else
percent_change = (total_yearly / open_yearly)
End If
volume_counter = volume_counter + ws.Cells(i, 7).Value
-----------------------------------------------------------

Coding step 2: The second part of the coding was much easier using a similiar line of code from the first part ( yearlastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row)
All i needed to do was write the formula and set the for loop to color the cells.
SAMPLE:
-------------------------
 If ws.Cells(j, 10).Value >= 0 Then
 ws.Cells(j, 10).Interior.ColorIndex = 4
 Else
 ------------------------
 
 Coding step 3: The third part uses the same code as the first twon steps with different variables.
 SAMPLE:
 -----------------
 If percentmax < ws.Cells(x, 11).Value Then
percentmax = ws.Cells(x, 11).Value
ws.Cells(2, 17).Value = percentmax
ws.Cells(2, 17).Style = "Percent"
ws.Cells(2, 16).Value = ws.Cells(x, 9).Value
ElseIf percentmin > ws.Cells(x, 11).Value Then
percentmin = ws.Cells(x, 11).Value
ws.Cells(3, 17).Value = percentmin
ws.Cells(3, 17).Style = "Percent"
ws.Cells(3, 16).Value = ws.Cells(x, 9).Value
End If
Next x
-----------------
