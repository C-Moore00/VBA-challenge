Attribute VB_Name = "StockReportCM"
Sub StockReport()
    'dim variables
    Dim row As Long
    Dim tick As String
    Dim volume As LongLong
    Dim dif As Double
    Dim openv As Double
    Dim closev As Double
    Dim tick2 As Integer
    Dim max As Double
    Dim min As Double
    Dim maxvol As LongLong
    Dim voltick As String
    Dim maxtick As String
    Dim mintick As String
    Dim sheet As Integer
    
    sheet = 0
    
    Do While sheet < Worksheets.Count
        Worksheets(sheet + 1).Activate
        'set variables for starting
        row = 2
        tick2 = 1
        max = 0
        min = 0
        maxvol = 0
    
        'create table headers
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"
        
        Do While Not Cells(row, 1) = "" 'as long as there are rows to check, continue
            If Cells(row, 1).Value = tick Then 'if ticker is still same, tally volume
                volume = volume + Cells(row, 7).Value
            Else 'else, set tick, open and volume to current row values
                tick = Cells(row, 1).Value
                openv = Cells(row, 3).Value
                volume = Cells(row, 7).Value
            End If
            If Not Cells(row + 1, 1) = tick Then 'if the next row isnt the same, then this is last row for given ticker and thus add values to sheet, tick2++, etc
                closev = Cells(row, 6).Value
                dif = closev - openv
                Cells(1 + tick2, 9) = tick
                Cells(1 + tick2, 10) = dif
                If dif < 0 Then
                    Cells(1 + tick2, 10).Interior.ColorIndex = 3
                Else
                    Cells(1 + tick2, 10).Interior.ColorIndex = 4
                End If
                Cells(1 + tick2, 10).NumberFormat = "0.00"
                Cells(1 + tick2, 11) = dif / openv
                Cells(1 + tick2, 11).NumberFormat = "0.00%"
                Cells(1 + tick2, 12) = volume
                Cells(1 + tick2, 12).NumberFormat = "0"
                tick2 = tick2 + 1
                If (dif / openv) < min Then
                    min = (dif / openv)
                    mintick = tick
                End If
                If (dif / openv) > max Then
                    max = (dif / openv)
                    maxtick = tick
                End If
                If volume > maxvol Then
                    maxvol = volume
                    voltick = tick
                End If
            End If
            row = row + 1 'tick row to next row
        Loop
        'create min / max / highvol fields and fill
        Cells(1, 16) = "Ticker"
        Cells(1, 17) = "Value"
        Cells(2, 15) = "Greatest % Increase"
        Cells(3, 15) = "Greatest & Decrease"
        Cells(4, 15) = "Greatest Total Volume"
        Cells(2, 16) = maxtick
        Cells(2, 17) = max
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 16) = mintick
        Cells(3, 17) = min
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(4, 16) = voltick
        Cells(4, 17) = maxvol
        Cells(4, 17).NumberFormat = "0"
        Range("O1:Q4").Columns.AutoFit
        Range("I1:L1").Columns.AutoFit
        sheet = sheet + 1
    Loop
End Sub

