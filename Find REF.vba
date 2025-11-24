Sub FindREFErrors()

    Dim ws As Worksheet
    Dim c As Range
    Dim nm As Name
    Dim co As ChartObject
    Dim ch As Chart
    Dim srs As Series
    Dim f As String
    Dim refFound As Boolean
    
    refFound = False
    
    Debug.Print String(60, "-")
    Debug.Print "Scanning workbook for #REF! errors: " & ThisWorkbook.Name
    Debug.Print "Time: " & Now
    Debug.Print String(60, "-")
    
    Application.ScreenUpdating = False
    
    '=== 1) Cells ===
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        For Each c In ws.UsedRange
            If c.HasFormula Then
                f = c.Formula
                If InStr(1, f, "#REF!", vbTextCompare) > 0 Then
                    Debug.Print "CELL REF ERROR: " & ws.Name & "!" & c.Address & " -> " & f
                    refFound = True
                End If
            End If
        Next c
        On Error GoTo 0
    Next ws
    
    '=== 2) Named Ranges ===
    For Each nm In ThisWorkbook.Names
        On Error Resume Next
        f = nm.RefersTo
        If InStr(1, f, "#REF!", vbTextCompare) > 0 Then
            Debug.Print "NAME REF ERROR: " & nm.Name & " -> " & f
            refFound = True
        End If
        On Error GoTo 0
    Next nm
    
    '=== 3) Chart Series ===
    For Each ws In ThisWorkbook.Worksheets
        For Each co In ws.ChartObjects
            Set ch = co.Chart
            On Error Resume Next
            For Each srs In ch.SeriesCollection
                f = srs.Formula
                If InStr(1, f, "#REF!", vbTextCompare) > 0 Then
                    Debug.Print "CHART REF ERROR: " & ws.Name & _
                                " - " & co.Name & " - Series: " & srs.Name & " -> " & f
                    refFound = True
                End If
            Next srs
            On Error GoTo 0
        Next co
    Next ws
    
    '=== 4) Data Validation ===
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        For Each c In ws.UsedRange
            If c.Validation.Type <> 0 Then
                f = c.Validation.Formula1
                If InStr(1, f, "#REF!", vbTextCompare) > 0 Then
                    Debug.Print "DATA VALIDATION REF ERROR: " & ws.Name & "!" & c.Address & " -> " & f
                    refFound = True
                End If
            End If
        Next c
        On Error GoTo 0
    Next ws
    
    '=== 5) Conditional Formatting ===
    Dim cf As FormatCondition
    For Each ws In ThisWorkbook.Worksheets
        For Each cf In ws.Cells.FormatConditions
            On Error Resume Next
            f = ""
            f = cf.Formula1
            If InStr(1, f, "#REF!", vbTextCompare) > 0 Then
                Debug.Print "COND FORMATTING REF ERROR: " & ws.Name & " -> " & f
                refFound = True
            End If
            On Error GoTo 0
        Next cf
    Next ws
    
    Application.ScreenUpdating = True
    
    If Not refFound Then
        Debug.Print "No #REF! errors found in the workbook."
    End If
    
    Debug.Print String(60, "-")
    Debug.Print

End Sub
