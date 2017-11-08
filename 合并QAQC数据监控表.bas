Attribute VB_Name = "模块1"
Sub 合并于总表() '将多个工作簿下的工作表依次对应合并到本工作簿下的工作表，即第一张工作表对应合并到第一张，第二张对应合并到第二张……
On Error Resume Next
Dim x As Variant, x1 As Variant, w As Workbook, wsh As Worksheet
Dim t As Workbook, ts As Worksheet, i As Integer, l As Integer, h As Long
Dim rng As Range, rng1 As Range
Dim color, start, xNum   As Integer

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Workbooks.Add
x = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", _
       Title:="Excel选择", MultiSelect:=True)
Set t = ActiveWorkbook
xNum = 1
For Each x1 In x

                If x1 <> False Then
                     Set w = Workbooks.Open(x1)
                          For i = 1 To w.Sheets.Count
                                If i > t.Sheets.Count Then t.Sheets.Add After:=t.Sheets(t.Sheets.Count)
                                t.Sheets(i).Name = w.Sheets(i).Name
                                 Set ts = t.Sheets(i)
                                 Set wsh = w.Sheets(i)
                                 l = ts.UsedRange.SpecialCells(xlCellTypeLastCell).Column
                                 h = ts.UsedRange.SpecialCells(xlCellTypeLastCell).Row
                                 Debug.Print "   lastrow:   " & h
                                      If xNum = 1 Then
                                                '  If l = 1 And h = 1 And ts.Cells(1, 1) = "" Then
                                                   wsh.Rows("1:" & wsh.Range("g7").End(xlDown).Row).Copy ts.Cells(1, 1)
                                                '   Else
                                                  ' wsh.Rows("1:" & wsh.Range("g7").End(xlDown).Row).Copy ts.Cells(1 + h, 1)
                                                  'End If
                                        Else
                                                 start = 3
                                                 Set rng = ts.Range("a3:a9")
                                                 For Each rng1 In rng
                                                     color = rng1.Interior.ColorIndex
                                                              If color <> -4142 Then
                                                                start = start + 1
                                                                Debug.Print "       loop：" & "book: " & x1 & "sheet:" & i & "   " & "startrow: " & start
                                                              Else
                                                                Exit For
                                                              End If
                                                   Next
                                                  'Debug.Print "合并总表：" & "book: " & x1 & "sheet:" & i & "   " & "startrow: " & start
                                                  wsh.Rows(start & ":" & wsh.Range("g7").End(xlDown).Row).Copy ts.Cells(1 + h, 1)
                                        
                                        End If
                                        
                         Debug.Print "合并总表：" & "book: " & x1 & "sheet:" & i & "   " & "startrow: " & start & Chr(10)
                         Next
                 w.Close
                 xNum = xNum + 1
                End If
Next
t.SaveAs Filename:="D:\合并表.xlsx", FileFormat:= _
xlOpenXMLWorkbook, CreateBackup:=False
    Sheets(1).Select
    Range("I8").Select
    
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "合并完成，保存于D盘，名称为：合并表.xlsx"
End Sub

Sub CLR()
Dim rng, rng1 As Range, start, color As Integer

h = ActiveWorkbook.Sheets(1).[a65536].End(xlUp).Row
                                         Debug.Print h


End Sub
Sub 还原空表()
'
' 还原空表 Macro
'

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Cells.Clear
    '清空当前表
Set t = ActiveWorkbook
Max = t.Sheets.Count
Debug.Print Max
 For i = 2 To Max
   t.Sheets(2).Delete
 Next
Range("F8").Select
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub


