Attribute VB_Name = "Module1"
Public Function Leq_btw_Timestamps(sheetName As String, startTime As Date, endTime As Date, data_column_no As Long)
' the date and time stamp must be in the first column, starting from 1
' sheetName: is the name of the sheet
' startTime and endTime must be date and time
' data_column_no: the first column is 1, then count forwards

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim total As Double
    Dim count As Long
    Dim currentTime As Variant
    Dim currentValue As Variant
    
    Set ws = ThisWorkbook.Sheets(sheetName) ' need to change this value
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    total = 0
    count = 0
    
    For i = 2 To lastRow  'Assuming headers in row 1
    
        currentTime = ws.Cells(i, 1).Value
        currentValue = ws.Cells(i, data_column_no).Value
        
        If IsDate(currentTime) And IsNumeric(currentValue) Then
            If currentTime >= startTime And currentTime <= endTime Then
                total = total + Application.WorksheetFunction.Power(10, currentValue / 10)
                count = count + 1
            End If
            End If
    Next i
    
    If count > 0 Then
        Leq_btw_Timestamps = 10 * (Application.WorksheetFunction.Log(total / count) / Application.WorksheetFunction.Log(10))
    Else
        Leq_btw_Timestamps = 0 ' Or return an error value
        
    End If
    
End Function

Public Function count_btw_timestamp(sheetName As String, startTime As Date, endTime As Date)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim total As Double
    Dim count As Long
    Dim currentTime As Variant
    
    Set ws = ThisWorkbook.Sheets(sheetName) ' need to change this value
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    total = 0
    count = 0
    
    For i = 2 To lastRow  'Assuming headers in row 1
    
        currentTime = ws.Cells(i, 1).Value
        
        If IsDate(currentTime) And IsNumeric(currentValue) Then
            If currentTime >= startTime And currentTime <= endTime Then
                count = count + 1
            End If
            End If
    Next i
    
    If count > 0 Then
        count_btw_timestamp = count
    Else
        count_btw_timestamp = 0 ' Or return an error value
        
    End If
End Function
