Attribute VB_Name = "合并多个工作簿到一个"
Sub 多工作簿合并() '将多个工作簿下的工作表依次对应合并到本工作簿下的工作表，即第一张工作表对应合并到第一张，第二张对应合并到第二张……
On Error Resume Next
Dim x As Variant, x1 As Variant, w As Workbook, wsh As Worksheet
Dim t As Workbook, ts As Worksheet, i As Integer, l As Integer, h As Long
Application.ScreenUpdating = False
Application.DisplayAlerts = False
x = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", _
       Title:="Excel选择", MultiSelect:=True)
Set t = ActiveWorkbook
For Each x1 In x
  If x1 <> False Then
  Set w = Workbooks.Open(x1)
  xNum = 1
    For i = 1 To w.Sheets.Count
        '3
        If i > t.Sheets.Count Then t.Sheets.Add After:=t.Sheets(t.Sheets.Count)
          Set ts = t.Sheets(i)  'gai
          t.Sheets(i).Name = w.Sheets(i).Name
          Set wsh = w.Sheets(i)
          l = ts.UsedRange.SpecialCells(xlCellTypeLastCell).Column
          h = ts.UsedRange.SpecialCells(xlCellTypeLastCell).Row
          '2
           If xNum = 1 Then
            wsh.Rows("1:" & wsh.Range("g7").End(xlDown).Row).Copy ts.Cells(1, 1)
           Else
             '1
             If l = 1 And h = 1 And ts.Cells(1, 1) = "" Then
                
                If i = 4 Then
                wsh.Rows("5:" & wsh.Range("g7").End(xlDown).Row).Copy ts.Cells(1, 1)
                ElseIf i = 6 Then
                wsh.Rows("7:" & wsh.Range("g7").End(xlDown).Row).Copy ts.Cells(1, 1)
                Else
                wsh.Rows("6:" & wsh.Range("g7").End(xlDown).Row).Copy ts.Cells(1, 1)
                End If
             
             
              Else
                If i = 4 Then
                wsh.Rows("5:" & wsh.Range("g7").End(xlDown).Row).Copy ts.Cells(h + 1, 1)
                ElseIf i = 6 Then
                wsh.Rows("7:" & wsh.Range("g7").End(xlDown).Row).Copy ts.Cells(h + 1, 1)
                Else
                wsh.Rows("6:" & wsh.Range("g7").End(xlDown).Row).Copy ts.Cells(h + 1, 1)
             
             End If
             '1end
          Cells.Select
          Selection.RowHeight = 13.5
          End If
          '2
        End If
        '3
     Next
     
     w.Close
     xNum = xNum + 1
    End If
Next
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub



