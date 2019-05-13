Sub UpdateLinks()
    Dim ExcelFile
    Dim exl As Object
    Set exl = CreateObject("Excel.Application")
     
    ExcelFile = "\\b8daprd\dfs\ssrhples\ASSISTENCIA EDUCACIONAL\MBA\Avaliação - Dezembro-2012\dados-MBA-UFSCAR.xlsx"
     
    Dim i As Integer
    Dim k As Integer
    With ActivePresentation.Slides(4)
        For k = 1 To .Shapes.Count
            On Error Resume Next
            .Shapes(k).LinkFormat.SourceFullName = ExcelFile
            If .Shapes(k).LinkFormat.SourceFullName = ExcelFile Then
                .Shapes(k).LinkFormat.AutoUpdate = ppUpdateOptionManual
            End If
        Next k
    End With
End Sub