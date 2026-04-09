Sub GuardarComoXlsx()
    Dim newPath As String
    newPath = Replace(ActiveWorkbook.FullName, ".xlsm", ".xlsx")
    ActiveWorkbook.SaveAs Filename:=newPath, FileFormat:=51
End Sub
