Sub ConditionalWrapAndFitColumns()
    Dim ws As Worksheet
    Dim col As Range, c As Range
    Dim maxLen As Long, cellLen As Long, cutoff As Long
    Dim cutoff As Integer, padding As Integer
    
    Set ws = ActiveSheet
    
    ' Set space padding & cutoff string length for "outlier" threshold (ignore anything bigger than this)
    cutoff = 80   ' <-- tweak this as you like
    padding = 2   ' <-- tweak this as you like
    
    ' Loop through each non-empty column
    For Each col In ws.UsedRange.Columns
        maxLen = 0
        
        ' Firstly, wrap all text in column
        col.WrapText = True
        
        ' Secondly, record longest cell length in column
        For Each c In col.Cells
            If Len(c.Value) > 0 And Len(c.Value) < cutoff Then
                cellLen = Len(c.Value)
                If cellLen > maxLen Then maxLen = cellLen
            End If
        Next c
        
        ' Thirdly, apply column width equal to largest cell plus an amount of padding (roughly 1 char = 1 unit width)
        If maxLen > 0 Then
            col.ColumnWidth = maxLen + padding
        End If
        
        ' Autofit row height (so wrapped text shows fully)
        col.EntireRow.AutoFit
    Next col
End Sub

