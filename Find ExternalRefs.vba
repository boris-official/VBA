Sub FindExternalLinks()
    Dim ws As Worksheet
    Dim c As Range
    Dim nm As Name
    Dim co As ChartObject
    Dim ch As Chart
    Dim srs As Series
    Dim f As String
    
    Dim extFound As Boolean
    extFound = False
    
    Debug.Print String(60, "-")
    Debug.Print "Scanning workbook: " & ThisWorkbook.Name
    Debug.Print "Time: " & Now
    Debug.Print String(60, "-")
    
    Application.ScreenUpdating = False
    
    '=== 1) Cells / formulas on each worksheet ===
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        For Each c In ws.UsedRange
            If c.HasFormula Then
                f = c.Formula
                ' External refs usually contain "[" (other workbook) or "http"
                If InStr(1, f, "[", vbTextCompare) > 0 Or _
                   InStr(1, f, "http", vbTextCompare) > 0 Then
                       
                    Debug.Print "CELL: " & ws.Name & "!" & c.Address & " -> " & f
                    extFound = True
                End If
            End If
        Next c
        On Error GoTo 0
    Next ws
    
    '=== 2) Named ranges ===
    For Each nm In ThisWorkbook.Names
        On Error Resume Next
        f = nm.RefersTo
        On Error GoTo 0
        
        If Len(f) > 0 Then
            If InStr(1, f, "[", vbTextCompare) > 0 Or _
               InStr(1, f, "http", vbTextCompare) > 0 Then
                   
                Debug.Print "NAME: " & nm.Name & " -> " & f
                extFound = True
            End If
        End If
    Next nm
    
    '=== 3) Charts (series formulas) ===
    For Each ws In ThisWorkbook.Worksheets
        For Each co In ws.ChartObjects
            Set ch = co.Chart
            On Error Resume Next
            For Each srs In ch.SeriesCollection
                f = srs.Formula
                If InStr(1, f, "[", vbTextCompare) > 0 Or _
                   InStr(1, f, "http", vbTextCompare) > 0 Then
                       
                    Debug.Print "CHART: " & ws.Name & " - " & co.Name & _
                                " - Series """ & srs.Name & """ -> " & f
                    extFound = True
                End If
            Next srs
            On Error GoTo 0
        Next co
    Next ws
    
    Application.ScreenUpdating = True
    
    If Not extFound Then
        Debug.Print "No external links found (cells, names, or chart series)."
    End If
    
    Debug.Print String(60, "-")
    Debug.Print
End Sub
