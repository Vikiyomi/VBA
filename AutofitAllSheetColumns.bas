Sub AutofitColumns()
f = ActiveSheet.Name
    Dim wrksht As Worksheet
    For Each wrksht In Worksheets
       
    wrksht.Select
    Cells.EntireColumn.AutoFit
    Cells.Font.Name = "Times New Roman"
    Cells.Font.Size = 12
    
    Next wrksht
    
Sheets(f).Select
End Sub
