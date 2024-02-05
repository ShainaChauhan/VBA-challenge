Attribute VB_Name = "Module1"
Sub Alfabetical_testing()

Dim ws As Worksheet
Dim sort_range As Range
Dim lastrow As Long 
Dim ticker As String 
Dim ychange As Double 
Dim Perc_chg As Double 
Dim stk_total As Double 
Dim sum_tbl_row As Long 
Dim start As Double 
Dim i As Integer

For Each ws In ThisWorkbook.Worksheets
ws.Activate

    start = 2 
    stk_total = 0
    sum_tbl_row = 2 
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row

    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"

    For i = 2 To lastrow
        ticker = Cells(start, "A").Value
    
        If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then            
            ychange = Cells(i, 6).Value - Cells(start, 3)
            Perc_chg = ychange / Cells(start, 3)
            stk_total = stk_total + Cells(i, "G").Value
            
            
            
            Range("I" & sum_tbl_row).Value = ticker
            Range("J" & sum_tbl_row).Value = ychange
            Range("K" & sum_tbl_row).Value = Format(Perc_chg, "0.00%")
            Range("L" & sum_tbl_row).Value = Format(stk_total, "$0,00")
            
            
            If Cells(sum_tbl_row, "J").Value < 0 Then
            
                Cells(sum_tbl_row, "J").Interior.ColorIndex = 3
            
            Else
            
                Cells(sum_tbl_row, "J").Interior.ColorIndex = 4
            
            End If

            
            stk_total = 0
                        
               
            
            sum_tbl_row = sum_tbl_row + 1
            start = i + 1
        
            
        Else 

            stk_total = stk_total + Cells(i, "G").Value

        End If

    Next i
    Columns("I:L").AutoFit
    Dim max As Double
    Dim min As Double
    Dim max_total As Double

    lastrow = Cells(Rows.Count, "J").End(xlUp).Row 'Resets lastrow for last range
    max = 0
    min = 0
    max_total = 0
    
    For i = 2 To lastrow

        If Cells(i, "K").Value > max Then

            max = Cells(i, "K").Value
            Cells(2, "O").Value = Cells(i, "I").Value

        ElseIf Cells(i, "K").Value < min Then

            min = Cells(i, "K").Value
            Cells(3, "O").Value = Cells(i, "I").Value

        End If

        If Cells(i, "L").Value > max_total Then
        
            max_total = Cells(i, "L").Value
            Cells(4, "O").Value = Cells(i, "I").Value
            
        End If
        
    Next i
    Cells(1, "O").Value = "Ticker"
    Cells(1, "P").Value = "Value"
    Cells(2, "N").Value = "Gratest % Increase"
    Cells(3, "N").Value = "Gratest % Decrease"
    Cells(4, "N").Value = "Gratest Total Volume"
    Cells(2, "P").Value = Format(max, "0.00%")
    Cells(3, "P").Value = Format(min, "0.00%")
    Cells(4, "P").Value = Format(max_total, "$0.00")

Columns("N:P").AutoFit
Next ws
End Sub
