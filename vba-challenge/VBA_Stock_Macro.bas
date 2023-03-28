Attribute VB_Name = "Module1"
Sub stock()

    'Variable Initialization
    Dim curSheet As Integer
    Dim curTick As String
    Dim curVol As Double
    Dim curStart As Double
    Dim curEnd As Double
    Dim tickNum As Integer
    
    'Loop Through Sheets
    For i = 1 To Application.Sheets.Count
    
        'Select Next Sheet
        Worksheets(i).Select
        
        'Set Headings
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'Format Columns
        Columns(11).NumberFormat = "[Green]0.00%;[Red]-0.00%"
        Columns(12).NumberFormat = "#,###,###,###"
        Columns(10).ColumnWidth = 15
        Columns(11).ColumnWidth = 15
        Columns(12).ColumnWidth = 20
        'Assign Initial Values
        curTick = Cells(2, 1).Value
        curStart = Cells(2, 3).Value
        tickNum = 1
        
        'Iterate through rows
        For j = 2 To Range("A1").End(xlDown).Row
            If (Cells(j, 1).Value = curTick) Then
                curVol = (curVol + Cells(j, 7).Value)
            Else
                curEnd = Cells((j - 1), 6).Value
                
                'Place values in table
                Cells((tickNum + 1), 9).Value = curTick
                Cells((tickNum + 1), 10).Value = (curEnd - curStart)
                
                'Ensure not dividing by 0
                If curStart - curEnd <> 0 Then
                    'Cells((tickNum + 1), 11).Value = (curStart / curEnd) / 100
                    Cells((tickNum + 1), 11).Value = (curEnd - curStart) / curStart
                Else
                    Cells((tickNum + 1), 11).Value = 0
                End If
                Cells((tickNum + 1), 12).Value = curVol
                
                'Reset Values
                tickNum = tickNum + 1
                curTick = Cells(j, 1).Value
                curStart = Cells(j, 3).Value
                curVol = 0
            End If
        Next j
        
        'Additional Functionality
        
        'Variable Initialization
        Dim maxPerc As Double
        Dim minPerc As Double
        Dim maxVol As Double
        Dim maxTick As String
        Dim minTick As String
        Dim volTick As String
        
        'Labels
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Columns(15).ColumnWidth = 20
        
        'Find & Set Values
        For j = 2 To Range("I1").End(xlDown).Row
        
            'Check Percentages
            If Cells(j, 11).Value > maxPerc Then
                maxPerc = Cells(j, 11).Value
                maxTick = Cells(j, 9).Value
            ElseIf Cells(j, 11).Value < minPerc Then
                minPerc = Cells(j, 11).Value
                minTick = Cells(j, 9).Value
            End If
            
            'Check Volume
            If Cells(j, 12).Value > maxVol Then
                maxVol = Cells(j, 12).Value
                volTick = Cells(j, 9).Value
            End If
        Next j
        
        'Set Values
        Cells(2, 16).Value = maxTick
        Cells(2, 17).Value = maxPerc
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 16).Value = minTick
        Cells(3, 17).Value = minPerc
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(4, 16).Value = volTick
        Cells(4, 17).Value = maxVol
        Cells(4, 17).NumberFormat = "#,###,###,###"
        
        'Reset Values
        maxPerc = 0
        minPerc = 0
        maxVol = 0
    Next i

End Sub
