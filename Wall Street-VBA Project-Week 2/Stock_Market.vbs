Attribute VB_Name = "Module1"
Sub StockMarket()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    
    ' --Variable Definitions--
    
    ' Ticker # Determination Variables
    
        Dim Ticker_Names As Object
        Dim erow As Long
        Dim Stock_Ticker As Range
        Dim x As Integer
    
    ' Loop Variables
    
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
    
    ' Table Variables
    
        Dim ticker As String
        Dim startprice As Double
        Dim iVal As Integer
        Dim endprice As Double
        Dim iRange As Range
        Dim row As Long
        Dim tablerow As Integer
        Dim volume As Double
    
    ' Bonus Variables
    
        Dim PercentRange As Range
        Dim MaxPercent As Double
        Dim MinPercent As Double
        Dim VolumeRange As Range
        Dim MaxVolume As Double
        
    ' Formatting Variable
    
        Dim lastrow As Long
     
    
    ' --Table & Bonus Headers--
    
    ' Table Headers
    
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(1, 17) = "Ticker"
        ws.Cells(1, 18) = "Value"
    
    ' Bonus Headers
    
        ws.Cells(2, 16) = "Greatest % Increase"
        ws.Cells(3, 16) = "Greatest % Decrease"
        ws.Cells(4, 16) = "Greatest Total Volume"
        
    
    ' --Initial Variable Assignments--
    
        row = 2
        tablerow = 2
        volume = 0
        
    
    ' --Ticker # Determination Code--
    
        Set Ticker_Names = CreateObject("Scripting.Dictionary")
        erow = CLng(ActiveSheet.Cells(1, 1).CurrentRegion.Rows.Count + 1)
    
        For Each Stock_Ticker In ws.Range("A2:A" & erow)
            If Not Ticker_Names.Exists(Stock_Ticker.Value) Then Ticker_Names.Add Stock_Ticker.Value, Nothing
        Next
    
        x = Ticker_Names.Count
        
    
    ' --Table Population Code--
    
        For i = 1 To x
            If i = x Then
                Exit For
            Else
                ticker = ws.Cells(row, 1).Value
                startprice = ws.Cells(row, 3).Value
                ws.Cells(tablerow, 9) = ticker
                iVal = Application.WorksheetFunction.CountIf(ws.Range("A2").EntireColumn, ticker)
                endprice = ws.Cells(iVal + row - 1, 6).Value
                ws.Cells(tablerow, 10) = endprice - startprice
                If startprice = 0 Or endprice = 0 Then
                    ws.Cells(tablerow, 11) = 0
                Else
                    ws.Cells(tablerow, 11) = (endprice / startprice) - 1
                End If
                Set iRange = ws.Range(ws.Cells(row, 7), ws.Cells(iVal + row - 1, 7))
                volume = Application.WorksheetFunction.Sum(iRange)
                ws.Cells(tablerow, 12) = volume
                row = row + iVal
                tablerow = tablerow + 1
                volume = 0
            End If
        Next i
        
        
    ' --Yearly Change Conditional Formatting--
        
        tablerow = tablerow - 1
        
        For j = 2 To tablerow
            If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 6
            End If
         Next j
      
    
    ' --Bonus Population Code--
    
    ' -Greatest % Increase & Decrease Code-
    
    ' Value Determination Code
    
        Set PercentRange = ws.Range(ws.Cells(2, 11), ws.Cells(x, 11))
        MaxPercent = Application.WorksheetFunction.Max(PercentRange)
        ws.Cells(2, 18).Value = MaxPercent
        MinPercent = Application.WorksheetFunction.Min(PercentRange)
        ws.Cells(3, 18).Value = MinPercent
    
    ' Ticker Determination Code
    
        For k = 2 To x
            If ws.Cells(k, 11).Value = MaxPercent Then
                ws.Cells(2, 17).Value = ws.Cells(k, 9).Value
            Else
            End If
            If ws.Cells(k, 11).Value = MinPercent Then
                ws.Cells(3, 17).Value = ws.Cells(k, 9).Value
            Else
            End If
        Next k
    
    ' -Greatest Total Volume Code-
    
    ' Value Determination Code
    
        Set VolumeRange = ws.Range(ws.Cells(2, 12), ws.Cells(x, 12))
        MaxVolume = Application.WorksheetFunction.Max(VolumeRange)
        ws.Cells(4, 18).Value = MaxVolume
    
    ' Ticker Determination Code
    
        For l = 2 To x
            If ws.Cells(l, 12).Value = MaxVolume Then
                ws.Cells(4, 17).Value = ws.Cells(l, 9).Value
            Else
            End If
        Next l
        
        
    ' --Worksheet Formatting--
    
    ' Columns
    
        ws.Columns("A:R").AutoFit
        ws.Columns("A:R").HorizontalAlignment = xlCenter
        ws.Columns("C:F" & "J").NumberFormat = "#,##0.00"
        lastrow = CLng(ws.Range("A1").CurrentRegion.Rows.Count)
        ws.Range(ws.Cells(2, 7), ws.Cells(lastrow, 7)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(2, 12), ws.Cells(tablerow, 12)).NumberFormat = "#,##0"
        ws.Columns("K").NumberFormat = "#,##0.00%"
        
    ' Bonus Values
    
        ws.Range(ws.Cells(2, 18), ws.Cells(3, 18)).NumberFormat = "#,##0.00%"
        ws.Cells(4, 18).NumberFormat = "#,##0"
        
    Next ws
    
End Sub


