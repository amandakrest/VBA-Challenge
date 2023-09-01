Sub WorksheetLoop2()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

Range("K1") = "Ticker"
Range("L1") = "Yearly Difference"
Range("M1") = "Percentage Change"
Range("N1") = "Total Stock Volume"

Dim ticker As String

Dim total_volume As Double
total_volume = 0

Dim closing As Double
closing = 0
Dim opening As Double
opening = 0

Dim percent_change As Double
percent_change = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

For i = 2 To 753001

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

ticker = Cells(i, 1).Value
total_volume = total_volume + Cells(i, 7).Value
closing = closing + Cells(i, 6).Value
opening = opening + Cells(i, 3).Value
percent_change = (closing - opening) / opening

Range("K" & Summary_Table_Row).Value = ticker
Range("N" & Summary_Table_Row).Value = total_volume
Range("L" & Summary_Table_Row).Value = closing - opening
Range("M" & Summary_Table_Row).Value = FormatPercent(percent_change)

Summary_Table_Row = Summary_Table_Row + 1

total_volume = 0

Else

total_volume = total_volume + Cells(i, 7).Value
closing = closing + Cells(i, 6).Value
opening = opening + Cells(i, 3).Value


End If

Dim Percent_Decrease As Double
Dim Percentage_Increase As Double
Dim Ticker_Increase As String
Dim Ticker_Decrease As String

Range("M" & Summary_Table_Row).Value = percent_change

If Percentage_Increase < Percentage_Change Then
    Percentage_Increase = percent_change
    Ticker_Increase = ticker
    
    End If

If Percent_Decrease > Cells(i + 1, 12).Value Then
    Percent_Decrease = Cells(i + 1, 11).Value
    Ticker_Decrease = Cells(i + 1, 1).Value

Else
    Percent_Decrease = Cells(i + 1, 11).Value
    Ticker_Decrease = Cells(i, 1).Value
    
    
Range("R2") = Ticker_Increase
Range("R3") = Ticker_Decrease


End If



Dim color_code As Range
Set color_code = Range("M2:M3001")

For Each Cell In color_code
If Cell.Value > "0" Then
Cell.Interior.ColorIndex = 4
Else
Cell.Interior.ColorIndex = 3

End If

Next

Range("R1") = "Ticker"
Range("S1") = "Value"
Range("Q2") = "Greatest % Increase"
Range("Q3") = "Greated % Decrease"
Range("Q4") = "Greatest Total Volume"



Range("S2") = greatest_increase
Range("S3") = greatest_decrease
Range("S4") = greatest_volume


Next ws


End Sub

    
