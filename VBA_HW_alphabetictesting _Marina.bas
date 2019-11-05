Attribute VB_Name = "Module1"
Sub TestData()

For Each ws In Worksheets

Dim tickername As String
Dim yearlychange As Double
Dim percentagechange As Double
Dim totalstock As Double
Dim initialvalue As Double
Dim finalvalue As Double
Dim Summarytablerow As Integer
Dim lastrow As Double
Dim lastrowsummary As Double
Dim max_increase As Variant
Dim min_increase As Variant
Dim max_stockvolume As Variant

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percetage Change"
ws.Range("L1").Value = "Total Stock Volume"

Summarytablerow = 2
totalstock = 0

initialvalue = ws.Cells(2, 3).Value
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        tickername = ws.Cells(i, 1).Value
        totalstock = totalstock + ws.Cells(i, 7).Value
        finalvalue = ws.Cells(i, 6).Value
        yearlychange = finalvalue - initialvalue
        'corrects in case the initial value is 0
            If initialvalue <> 0 Then
                 percentagechange = yearlychange / initialvalue
             Else
                percentagechange = 0
            End If
        '--------------------------------
        'writes data on Summary Table and colo-codes the yearly change
        ws.Range("I" & Summarytablerow).Value = tickername
        ws.Range("J" & Summarytablerow).Value = yearlychange
        ws.Range("K" & Summarytablerow).Value = percentagechange
        ws.Range("L" & Summarytablerow).Value = totalstock
        'color the yearly change in green if >0 or red if = < 0
            If yearlychange > 0 Then
            ' Color green if positive grow
                 ws.Range("J" & Summarytablerow).Interior.ColorIndex = 4
            Else
            ' Color red if no grow or negative
                ws.Range("J" & Summarytablerow).Interior.ColorIndex = 3
            End If
        '------------------------------------
        
        totalstock = 0
        initialvalue = ws.Cells(i + 1, 3).Value
        Summarytablerow = Summarytablerow + 1
    
    Else
        totalstock = totalstock + ws.Cells(i, 7).Value
    End If

Next i
          
     'Corrects the format of Summary Table
        lastrowsummary = ws.Cells(Rows.Count, "I").End(xlUp).Row
        For k = 2 To lastrowsummary
            ws.Cells(k, 11).Style = "Percent"
            ws.Cells(k, 12).Style = "Currency"
        Next k
     '----------------------------------------
     'Calculate the greatest % increase
        max_increase = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrowsummary))
        min_increase = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrowsummary))
        max_stockvolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrowsummary))
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
      
        MsgBox (lastrowsummary)
        MsgBox (max_stockvolume)
        For x = 2 To lastrowsummary
            If ws.Cells(x, 12).Value = max_stockvolume Then
                ws.Cells(4, 14).Value = "Greatest Total Volume"
                ws.Cells(4, 15).Value = ws.Cells(x, 9).Value
                ws.Cells(4, 16).Value = max_stockvolume
                MsgBox (max_stockvolume)
            End If
          Next x
          
        For m = 2 To lastrowsummary
            If ws.Cells(m, 11).Value = max_increase Then
                ws.Cells(2, 14).Value = "Greatest % Increase"
                ws.Cells(2, 15).Value = ws.Cells(m, 9).Value
                ws.Cells(2, 16).Value = max_increase
                MsgBox (min_increase)
            ElseIf ws.Cells(m, 11).Value = min_increase Then
                ws.Cells(3, 14).Value = "Greatest % Decrease"
                ws.Cells(3, 15).Value = ws.Cells(m, 9).Value
                ws.Cells(3, 16).Value = min_increase
                MsgBox (min_increase)
            End If
        Next m
          
           'Correct the format
            ws.Range("P2:P3").Style = "Percent"
            ws.Cells(4, 16).Style = "Currency"
            
     
Next ws

End Sub


