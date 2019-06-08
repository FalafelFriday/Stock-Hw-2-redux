Attribute VB_Name = "Module1"
Sub multiyearstock()
'First the easy stuff set the values of those cells tht cant change
For Each ws In Worksheets
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 16).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 12).Value = "Total Stock Volume"
 ws.Cells(1, 17).Value = "Value"
 ws.Cells(2, 15).Value = "Greatest % Increase"
 ws.Cells(3, 15).Value = "Greatest % Decrease"
 ws.Cells(4, 15).Value = "Greatest Total Volume"
'Now we Dim out everything for step 1
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Long

Dim Stock_name As String

Dim open_yearly As Double

Dim volume_counter As Double
volume_counter = 0

Dim total_yearly As Double
total_yearly = 0
Dim percent_change As Double

Dim P_P As Long
P_P = 2
'P_P is a sortof placeholder that will help me later
Dim Table_row As Integer
Table_row = 2
'Set the For loop
'here we also use P_P to make sure we can get the yearly open for each stock

For i = 2 To lastrow
open_yearly = ws.Cells(P_P, 3).Value
'Formulas that will calculate total volume,percent change total yearly change and stock name
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Stock_name = ws.Cells(i, 1).Value
total_yearly = total_yearly + (Cells(i, 6).Value - open_yearly)
If open_yearly = 0 Then
percent_change = 0
Else
percent_change = (total_yearly / open_yearly)
End If
volume_counter = volume_counter + ws.Cells(i, 7).Value
'Now we adjust the ranges on our table to corespond to valuse we got from our formulas
ws.Range("I" & Table_row).Value = Stock_name
ws.Range("L" & Table_row).Value = volume_counter
ws.Range("J" & Table_row).Value = total_yearly
ws.Range("K" & Table_row).Value = percent_change
ws.Range("K" & Table_row).Style = "Percent"
'Reset
Table_row = Table_row + 1
volume_counter = 0
total_yearly = 0
P_P = i + 1

open_yearly = ws.Cells(P_P, 3).Value

Else
    volume_counter = volume_counter + ws.Cells(i, 7).Value
 End If
 Next i
 'Step 1 complete :) Step 2 color the cells
 'Dim out a new last row
 Dim yearlastrow As Long
 
 yearlastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
 'Create a new for loop
 For j = 2 To yearlastrow
 'Create a conditional to color the cells
 If ws.Cells(j, 10).Value >= 0 Then
 ws.Cells(j, 10).Interior.ColorIndex = 4
 Else
 ws.Cells(j, 10).Interior.ColorIndex = 3
 End If
 Next j
'Step 2 complete :) Next step the challenging part
'Dim out a new last row
Dim percentlastrow As Long
percentlastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
'DIm out variable for percent max and min
Dim percentmax As Double
percentmax = 0
Dim percentmin As Double
percentmin = 0
'New for loop
For x = 2 To percentlastrow
'Conditional Statement to find max and min
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
'Dim out final last row
Dim volumelastrow As Long
volumelastrow = ws.Cells(Rows.Count, 12).End(xlUp).Row
'Dim out a variable
Dim totalvolumemax As Double
totalvolumemax = 0
'Dim out the final for loop
For i = 2 To volumelastrow
'Conditional to solve for the variable
If totalvolumemax < ws.Cells(i, 12).Value Then
totalvolumemax = ws.Cells(i, 12).Value
ws.Cells(4, 17).Value = totalvolumemax
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
End If
Next i
 Next ws

End Sub

