Attribute VB_Name = "Module1"

' VBA Code: Auto-Generate Summary Report
' Counts: Total Rainy Days, Total Cold Days, Total High Rental Days

Sub CreateSummaryTable()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rainCount As Long, coldCount As Long, highRentalCount As Long
    Dim weatherCol As Integer, tempCol As Integer, countCol As Integer
    Dim i As Long

    ' Set your worksheet
    Set ws = ThisWorkbook.Sheets("Data")

    ' Find the last used row
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Find column numbers
    Dim headerCell As Range
    For Each headerCell In ws.Rows(1).Cells
        Select Case LCase(Trim(headerCell.Value))
            Case "weather"
                weatherCol = headerCell.Column
            Case "temp_real_c"
                tempCol = headerCell.Column
            Case "count"
                countCol = headerCell.Column
        End Select
    Next headerCell

    ' Loop through data
    For i = 2 To lastRow
        If LCase(Trim(ws.Cells(i, weatherCol).Value)) = "rain" Then
            rainCount = rainCount + 1
        End If
        If IsNumeric(ws.Cells(i, tempCol).Value) Then
            If ws.Cells(i, tempCol).Value < 10 Then
                coldCount = coldCount + 1
            End If
        End If
        If IsNumeric(ws.Cells(i, countCol).Value) Then
            If ws.Cells(i, countCol).Value > 1000 Then
                highRentalCount = highRentalCount + 1
            End If
        End If
    Next i

    ' Paste the summary below your data
    Dim summaryStartRow As Long
    summaryStartRow = lastRow + 3

    ws.Cells(summaryStartRow, 1).Value = "Summary Report"
    ws.Cells(summaryStartRow, 1).Font.Bold = True

    ws.Cells(summaryStartRow + 1, 1).Value = "Total Rainy Days"
    ws.Cells(summaryStartRow + 1, 2).Value = rainCount

    ws.Cells(summaryStartRow + 2, 1).Value = "Total Cold Days (Temp < 10¡C)"
    ws.Cells(summaryStartRow + 2, 2).Value = coldCount

    ws.Cells(summaryStartRow + 3, 1).Value = "Total High Rental Days (Rentals > 1000)"
    ws.Cells(summaryStartRow + 3, 2).Value = highRentalCount

    MsgBox "Summary Report Created Successfully!", vbInformation

End Sub

' Advanced Formatting (only highlight specific columns)
' Instead of coloring the whole row, only highlight weather/temp/count cells.

Sub SoftHighlightSpecificCells()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim weatherCol As Integer, tempCol As Integer, countCol As Integer

    Set ws = ThisWorkbook.Sheets("Data")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Find columns
    Dim headerCell As Range
    For Each headerCell In ws.Rows(1).Cells
        Select Case LCase(Trim(headerCell.Value))
            Case "weather"
                weatherCol = headerCell.Column
            Case "temp_real_c"
                tempCol = headerCell.Column
            Case "count"
                countCol = headerCell.Column
        End Select
    Next headerCell

    ' Highlight only specific cells
    For i = 2 To lastRow
        If LCase(Trim(ws.Cells(i, weatherCol).Value)) = "rain" Then
            ws.Cells(i, weatherCol).Interior.Color = RGB(255, 250, 205) ' Pastel Yellow
        End If
        
        If IsNumeric(ws.Cells(i, tempCol).Value) Then
            If ws.Cells(i, tempCol).Value < 10 Then
                ws.Cells(i, tempCol).Interior.Color = RGB(224, 255, 255) ' Pastel Blue
            End If
        End If
        
        If IsNumeric(ws.Cells(i, countCol).Value) Then
            If ws.Cells(i, countCol).Value > 1000 Then
                ws.Cells(i, countCol).Interior.Color = RGB(240, 255, 240) ' Pastel Green
            End If
        End If
    Next i

    MsgBox "Specific Cells Highlighted Successfully!", vbInformation

End Sub

' Create Pivot Table Automatically
' Creating a pivot table that shows total bikes rented by season and weather.


Sub CreatePivotTable()

    Dim ws As Worksheet
    Dim pvtWs As Worksheet
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim dataRange As Range
    Dim lastRow As Long
    Dim lastCol As Long

    ' Set source worksheet
    Set ws = ThisWorkbook.Sheets("Data")

    ' Find data range
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Create a new sheet for Pivot Table
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("PivotTable").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set pvtWs = ThisWorkbook.Sheets.Add
    pvtWs.Name = "PivotTable"

    ' Create Pivot Cache
    Set pvtCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    ' Create Pivot Table
    Set pvt = pvtCache.CreatePivotTable( _
        TableDestination:=pvtWs.Cells(1, 1), _
        TableName:="BikesPivot")

    ' Set fields
    With pvt
        .PivotFields("season").Orientation = xlRowField
        .PivotFields("season").Position = 1
        .PivotFields("weather").Orientation = xlRowField
        .PivotFields("weather").Position = 2
         .AddDataField .PivotFields("count"), "Total Rentals", xlSum
        .PivotFields("Total Rentals").NumberFormat = "#,##0"
    End With

    MsgBox "Pivot Table Created Successfully!", vbInformation

End Sub


