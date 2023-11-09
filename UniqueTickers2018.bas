Attribute VB_Name = "UniqueTickers2018"
Sub tickvalues()

Dim wb As Workbook
Dim ws As Worksheet
Dim RNG As Range
Dim lrow As Long


Set wb = ThisWorkbook
Set ws = wb.Worksheets("2018")
Set RNG = ws.Range("I1")


ws.Range("A1:A753001").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=RNG, unique:=True


' Use a range function to assign Cell Value and make Header as Bold Text
Range("I1").Value = ("Ticker")
Range("I1").Font.Bold = True
Range("J1").Value = ("Yearly Change")
Range("J1").Font.Bold = True
Range("K1").Value = ("PercentChange")
Range("K1").Font.Bold = True
Range("L1").Value = ("Total Stock Volume")
Range("L1").Font.Bold = True

'Create new text in cells for Chnages over time
Range("O3").Value = ("Greatest % Increase")
Range("O3").Font.Bold = True
Range("O4").Value = ("Greatest % Decrease")
Range("O4").Font.Bold = True
Range("O5").Value = ("Greatest Total Volume")
Range("O5").Font.Bold = True


End Sub
