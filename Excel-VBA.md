# 重命名带链接的单元格的文字

```
Sub RenameCells()

    Dim hl As Hyperlink
    
    ThisWorkbook.Sheets(1).Activate
    
    For Each hl In ActiveSheet.Hyperlinks
        hl.TextToDisplay = "查看"  '''不知道哪里设置的不对，有部分单元格未被更改
    Next
    
End Sub
```

# 修改带链接单元格的链接地址

```
Sub ModifyAdd()

    Dim records As Integer
    
    ThisWorkbook.Sheets(1).Activate
    
    records = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To records
        OriCellPos = "A" + CStr(i)
        OldAdd = Range(OriCellPos).Hyperlinks(1).Address
        NewAdd = Left(OldAdd, 7) + "\" + Mid(OldAdd, 12, 14)
        Range(OriCellPos).Hyperlinks(1).Address = NewAdd
    Next
    
End Sub
```
