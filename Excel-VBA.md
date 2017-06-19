Sub RenameCells()

    Dim hl As Hyperlink
    
    ThisWorkbook.Sheets(1).Activate
    
    For Each hl In ActiveSheet.Hyperlinks
        hl.TextToDisplay = "查看"  '''不知道哪里设置的不对，有部分单元格未被更改
    Next
    
End Sub
