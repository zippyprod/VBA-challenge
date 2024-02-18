Attribute VB_Name = "UniqueTickers2018"
Sub Multiple_Year_Stock()

Dim ws As Worksheet

'Loop through all of the worksheets
For Each ws In ThisWorkbook.Worksheets
ws.Activate


'Find the Last Row
LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
'Present a MsgBox (LastRow)
Location = 2

'Ticker Name
ws.Range("I1").Value = "Ticker Name"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"


Dim TickerName As String
Dim openvalue As Double
Dim closevalue As Double
Dim totalvolume As Double

openvalue = ws.Cells(2, 3)
closevalue = 0
totalvolume = 0

' Define Greatest Percent Increase

Dim greatestpercentchange As Double


' Ticker Name pulling unique
For i = 2 To LastRow


If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then

    closevalue = ws.Cells(i, 6).Value
    
    ws.Cells(Location, 9).Value = ws.Cells(i, 1).Value

    ws.Cells(Location, 10).Value = closevalue - openvalue
    
    ws.Cells(Location, 11).Value = (closevalue - openvalue) / openvalue

    openvalue = ws.Cells(i + 1, 3).Value

    closevalue = 0
    
    greatestpercentchange = ws.Cells(Location, 11).Value
    
    If greatestpercentchange > ws.Range("Q2").Value Then
    
    ws.Range("P2").Value = ws.Cells(Location, 9).Value
    
    ws.Range("Q2").Value = greatestpercentchange
    
    End If
    
' Find the Greatest Negative Change

    If greatestpercentchange < ws.Range("Q3").Value Then
    
    ws.Range("P3").Value = ws.Cells(Location, 9).Value
    
    ws.Range("Q3").Value = greatestpercentchange
    
    End If
    
'Find theGreatest Stock Volume
    
    If ws.Cells(Location, 12).Value > ws.Range("Q4").Value Then
    
    ws.Range("P4").Value = ws.Cells(Location, 9).Value
    
    ws.Range("Q4").Value = ws.Cells(Location, 12).Value
    
    End If

ws.Cells(2, 17).NumberFormat = "0.00%"

ws.Cells(3, 17).Value = Format(ws.Cells(3, 17).Value * 100, "0.00") & "%"

    
    ' Change cell colors
 
 If ws.Cells(Location, 10).Value >= 0 Then
ws.Cells(Location, 10).Interior.ColorIndex = 4
Else
ws.Cells(Location, 10).Interior.ColorIndex = 3

End If

    ' Change the Number to a Percentage
    
     ws.Cells(Location, 11).Value = Format(ws.Cells(Location, 11).Value * 100) & "%"
    
    ' Convert column j into number

     ws.Cells(Location, 10).Value = Format(ws.Cells(Location, 10).Value) & "0"
    
' Find the Total Volume

    totalvolume = totalvolume + ws.Cells(i, 7).Value
    ws.Cells(Location, 12).Value = totalvolume
    totalvolume = 0
    Location = Location + 1

Else
    totalvolume = totalvolume + ws.Cells(i, 7).Value
     
End If

Next i

Next ws

End Sub

