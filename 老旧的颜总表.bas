Attribute VB_Name = "模块2"
Sub 在报名表中插入空行()
Attribute 在报名表中插入空行.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
'

'
begin = ActiveCell.Row
ending = ActiveCell.End(xlDown).Value
For i = begin To ending + begin - 1
b = Range("a" & i + 1).Value
a = Range("a" & i).Value
Debug.Print "rownum:" & i
Debug.Print "row↑" & a
Debug.Print "row↓" & b
diff = b - a
 crtRow = i
   If diff > 1 Then
    For cishu = 1 To b - a - 1
    Debug.Print "差值:" & b - a - 1
   
    Rows(i + 1 & ":" & i + 1).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(i + 1, 1).Value = Cells(i, 1).Value + 1
    Cells(i + 1, 2) = "=INDIRECT(ADDRESS(" & crtRow & "," & cishu + 2 & "))"
    '=INDIRECT(ADDRESS(53,3))
    i = i + 1
    Next cishu
   End If
   
Next i
    Columns("B:B").Select
    Columns("B:B").Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
End Sub

Sub 处理颜总的接龙报名()
Attribute 处理颜总的接龙报名.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 处理颜总的接龙报名 宏
'

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Set tou = ActiveCell


    Range(tou, tou.End(xlDown)).Select
    Selection.Replace What:="，", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=".", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="．", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="（", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="、", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="。", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="，", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="：", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=":", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="已", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="费", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
For i = tou.Row To tou.End(xlDown).Row
Set rng1 = Cells(i, 1)
    strin = rng1.Text
  spcNum = InStr(2, rng1, " ")
    If InStr(4, rng1, "位") Then
        rng1 = Left(strin, spcNum) & Mid(strin, InStr(4, rng1, "位") + 1, Len(strin) - InStr(4, strin, "位"))
    End If
Next i
Range(tou, tou.End(xlDown)).Select
        'A列按空格数据分列
    Selection.TextToColumns Destination:=tou, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="(", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
        '关掉数据分列
    Selection.TextToColumns Destination:=tou, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :="", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
begin = tou.Row
ending = tou.End(xlDown).Value
For i = begin To begin + ending - 1
If i > begin Then
Cells(begin, i - begin + 3) = Cells(i, Cells(i, 2).End(xlToRight).Column).Text
End If
b = Range("a" & i + 1).Value
a = Range("a" & i).Value
diff = b - a
 crtRow = i
   If diff > 1 Then
    For cishu = 1 To b - a - 1
    Rows(i + 1 & ":" & i + 1).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(i + 1, 1).Value = Cells(i, 1).Value + 1
    Cells(i + 1, 2) = "=INDIRECT(ADDRESS(" & crtRow & "," & cishu + 2 & "))"
    i = i + 1
    Next cishu
   End If
   
Next i
Set name1 = Cells(tou.Row, 2)
Set name2 = Cells(tou.End(xlDown).Row, 2)
    Range(name1 & ":" & name2).Copy
    Range(name1 & ":" & name2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
For i = begin + 1 To begin + tou.End(xlDown).Value - 1
Cells(i, 3) = Cells(begin, i - begin + 3)
Next i
Set name1 = Cells(tou.Row, 4)
Set name2 = Cells(tou.End(xlDown).Row, 100)
Range(name1 & ":" & name2).ClearContents
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub 分列()
Attribute 分列.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 分列 宏
'

'按照空格分列，重复的分隔符视为一个
  '  Range("A11:A26").Select
    Range("A15:A100").TextToColumns Destination:=Range("A15"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :=".", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        
 '分列结束，关掉分裂功能
    Range("a2").TextToColumns Destination:=Range("A2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :=".", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
End Sub

Sub 填色宏()
Attribute 填色宏.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 单元格涂色
'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
Cells(3, 3).Interior.color = 4000
End Sub


Sub 清空个字()
Attribute 清空个字.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏16 宏
'

'
    Range("G17").Select
    Selection.ClearContents
End Sub
Sub SHIPXY()
Attribute SHIPXY.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' 经纬度
'

'
    ro = ActiveCell.Row
    co = ActiveCell.Column
    
    ActiveSheet.Paste
    Cells(ro + 2, co).Select
End Sub
Sub 删除并复制()
Attribute 删除并复制.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏22 宏
'

'
    Sheets("Sheet10").Cells.ClearContents
    Sheets("Sheet9").Range("A356:a483").Copy Sheets("Sheet10").Range("a1")
    Sheets("Sheet10").[a13].Select
End Sub
