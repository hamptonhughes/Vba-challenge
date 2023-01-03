Sub stock_Picker()

Dim ws As Worksheet
For Each ws In Worksheets



    'Define all the variables we will need
    
    Dim ticker As String
    Dim i As Double
    Dim rowcount As Double
    Dim Yearstart As Double
    Dim Yearend As Double
    Dim lastrow As Double
    Dim volume As Double
    Dim YoYchange As Double

    'Start Volume and Rowcount at 0 and 2
    
    volume = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    rowcount = 2

    'Set Headers for summary tables
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "YoY Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
    'Loop through all rows

    For i = 2 To lastrow

    'If first ticker doesn't equal previous row, start counting volume and assign that row as the year start
    
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        Yearstart = ws.Cells(i, 3).Value
        volume = ws.Cells(i, 7).Value + volume
    
        'If Ticker is same as row below, just keep summing volume
    
    
        ElseIf ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        volume = ws.Cells(i, 7).Value + volume
    
        'Once Ticker does not equal the next row, assign that row as year end and ticker, count volume for one more row.
        'Calculate YoY Change for that ticker
        'Calculate % change
    
        Else
        Yearend = ws.Cells(i, 6).Value
        volume = ws.Cells(i, 7).Value + volume
        ticker = ws.Cells(i, 1).Value
        YoYchange = Yearend - Yearstart
        Percentchange = YoYchange / Yearstart
    
        'Record ticker, volume, Yoy change, and percent change into a table in rows 9-12
    
        ws.Cells(rowcount, 9).Value = ticker
        ws.Cells(rowcount, 12).Value = volume
        ws.Cells(rowcount, 10).Value = YoYchange
        ws.Cells(rowcount, 11).Value = Percentchange
    
        
        'Add 1 to the running row count.  Reset volume to 0 for next ticker
        rowcount = rowcount + 1
        volume = 0
 
        End If

    
    Next i
    
    
'Set new variables for new For statement which will set the formatting
Dim k As Double
Dim lastrowsummary As Double

'Find the last row
lastrowsummary = ws.Cells(Rows.Count, 11).End(xlUp).Row

'For statement to format the cells in the summary table
For k = 2 To lastrowsummary

'Set format stye to percent for the Yoy percent change
ws.Cells(k, 11).Style = "Percent"

    'If statement to change color to green or red depending on if percent change was +/-
    If ws.Cells(k, 11) >= 0 Then
    
    ws.Cells(k, 11).Interior.ColorIndex = 4
    
    Else
    
    ws.Cells(k, 11).Interior.ColorIndex = 3
    
    End If

    
    If ws.Cells(k, 10) >= 0 Then
    
    ws.Cells(k, 10).Interior.ColorIndex = 4
    
    Else
    
    ws.Cells(k, 10).Interior.ColorIndex = 3
    
    End If

Next k


'Set Variables and definitions for second summary table

Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim maxVol As Double
Dim IncreaseTick As String
Dim DecreaseTick As String
Dim VolTick As String
Dim l As Integer
Dim m As Integer
Dim n As Integer

    'If statement comparing first two rows, assign the higher value as biggest % increase.  Assign ticker too.

    If ws.Cells(2, 11).Value < ws.Cells(3, 11).Value Then
    MaxDecrease = ws.Cells(2, 11).Value
    DecreaseTick = ws.Cells(2, 9).Value
    
    Else
    MaxDecrease = ws.Cells(3, 11).Value
    DecreaseTick = ws.Cells(3, 9).Value
   
    End If

    'For statement to compare each row against the max % increase variable.  Once finished, we will have highest % increase and the associated ticker
    For l = 4 To lastrowsummary
        If ws.Cells(l, 11).Value < MaxDecrease Then
        MaxDecrease = ws.Cells(l, 11).Value
        DecreaseTick = ws.Cells(l, 9).Value
   
        End If
    
    Next l

    'Same as the last set of scripts, except finding the biggest % decrease
    If ws.Cells(2, 11).Value > ws.Cells(3, 11).Value Then
    MaxIncrease = ws.Cells(2, 11).Value
    IncreaseTick = ws.Cells(2, 9).Value
    Else
    MaxIncrease = ws.Cells(3, 11).Value
    IncreaseTick = ws.Cells(3, 9).Value
    End If


    For m = 4 To lastrowsummary
        If ws.Cells(m, 11).Value > MaxIncrease Then
        MaxIncrease = ws.Cells(m, 11).Value
        IncreaseTick = ws.Cells(m, 9).Value
        End If
    
    Next m
    
    'same as last set of scripts, except finding max volume
    If ws.Cells(2, 12).Value > ws.Cells(3, 12).Value Then
    maxVol = ws.Cells(2, 12).Value
    VolTick = ws.Cells(2, 9).Value
    
    Else
    maxVol = ws.Cells(3, 12).Value
    VolTick = ws.Cells(3, 9).Value
    
    End If

    For n = 4 To lastrowsummary
        If ws.Cells(n, 12).Value > maxVol Then
        maxVol = ws.Cells(n, 12)
        VolTick = ws.Cells(n, 9).Value
        
        End If
    

    Next n

'Record Greatest increase/decrease, greatest volume into new table

ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(2, 15).Value = IncreaseTick
ws.Cells(3, 15).Value = DecreaseTick
ws.Cells(4, 15).Value = VolTick
ws.Cells(2, 16).Value = MaxIncrease
ws.Cells(3, 16).Value = MaxDecrease
ws.Cells(4, 16).Value = maxVol
ws.Range("P2:P3").Style = "Percent"


Next ws



End Sub