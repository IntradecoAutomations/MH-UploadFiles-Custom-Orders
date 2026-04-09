Sub UnhideAllSheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Drop Down" Then
            ws.Visible = xlSheetVisible
        End If
    Next ws
End Sub
