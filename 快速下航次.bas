Attribute VB_Name = "快速下航次"
Sub 快速下航次()
Attribute 快速下航次.VB_ProcData.VB_Invoke_Func = "Q\n14"
Dim kaishi, jieshu, i As Integer, str, abc As String

r = ActiveCell.Row
c = ActiveCell.Column

Application.ScreenUpdating = 0
Application.DisplayAlerts = 0
'首先统一航次命令格式

    Cells(r, 19).Select
    Selection.Replace What:="v", Replacement:="V", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(r, 19).Select
    Selection.Replace What:=Chr(10), Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(r, 2).Select
    Selection.Replace What:="v", Replacement:="V", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(r, 19).Select
    Selection.Replace What:="【", Replacement:="(", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(r, 19).Select
    Selection.Replace What:="】", Replacement:=")", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(r, 19).Select
    Selection.Replace What:="—", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(r, 19).Select
    Selection.Replace What:="--", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(r, 19).Select
    Selection.Replace What:="，", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(r, 19).Select
    Selection.Replace What:=",", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(r, 19).Select
    Selection.Replace What:="吨", Replacement:="MT", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(r, 19).Select
    Selection.Replace What:="（", Replacement:="(", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(r, 19).Select
    Selection.Replace What:="）", Replacement:=")", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(r, 19).Select
    Selection.Replace What:="±5%", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

str = Trim(Cells(r, 19).Text)
If str = "" Then
GoTo endsub:
End If

If Left(str, 1) <> "V" Then
str = "V" & str
'MsgBox "判断并添加开头V：" & str
End If
If Mid(str, 6, 1) = "&" Then
Cells(r, 22) = Right(str, Len(str) - 6)
str = Left(str, 5) & Right(str, Len(str) - 10)
End If
If Len(str) > 49 Then
Cells(r, 22) = Right(str, Len(str) - InStr(17, str, "V1") + 1)
str = Trim(Left(str, InStr(17, str, "V1") - 1))
End If
If Mid(str, 6, 1) = " " Then
str = Left(str, 5) & Right(str, Len(str) - 6)
End If
If Mid(str, 6, 1) <> "航" Then
str = Left(str, 5) & "航次命令" & Right(str, Len(str) - 5)
End If
If Mid(str, 10, 1) = " " Then
str = Left(str, 9) & Right(str, Len(str) - 10)
End If
If Mid(str, 10, 1) <> "(" Then
str = Left(str, 9) & "(" & Right(str, Len(str) - 9)
'MsgBox "加入(：" & str
End If
If InStr(9, Left(str, Len(str) - 2), ")") <> 0 Then
str = Left(str, InStr(9, str, ")") - 1) & Right(str, Len(str) - InStr(9, str, ")"))
'MsgBox "删除货量前)：" & str
End If
If Mid(str, InStr(13, str, "T") - 6, 1) <> " " Then
str = Left(str, InStr(13, str, "T") - 6) & " " & Right(str, Len(str) - InStr(13, str, "T") + 6)
'MsgBox "货量前加空格：" & str
End If
If Mid(str, InStr(13, str, "T") + 1, 1) <> " " Then
str = Left(str, InStr(13, str, "T")) & " " & Right(str, Len(str) - InStr(13, str, "T"))
'MsgBox "货种前加空格：" & str
End If
If Right(str, 1) = "令" Then
str = Left(str, Len(str) - 4)
'MsgBox "删掉最后的航次命令：" & str
End If
If Right(str, 1) <> ")" Then
str = str & ")"
'MsgBox "最后加)：" & str
End If
Cells(r, 19) = str
        
 '处理完毕
 
' MsgBox "处理完毕" & i
 
 '航次结束，更新船舶航行计划

str = Cells(r, 19).Text
'MsgBox "s" & i & ":" & Left(str, 5)
'MsgBox "b" & i & ":" & Range("b" & i).Text
'MsgBox Range("b" & i).Text = Left(str, 5)

kao = InStr(10, str, "(", 1) + 1
'MsgBox kao
lenkao = InStr(12, str, "-", 1) - InStr(10, str, "(", 1) - 1

xie = InStr(12, Cells(r, 19), "-", 1) + 1
'MsgBox xie
lenxie = InStr(16, Cells(r, 19), " ", 1) - InStr(13, Cells(r, 19), "-", 1) - 1
'MsgBox lenxie
cargo = InStr(23, Cells(r, 19), " ", 1) + 1
'MsgBox cargo

lencar = InStr(25, Cells(r, 19), ")", 1) - InStr(23, Cells(r, 19), " ", 1) - 1
'MsgBox lencar
quanti = InStr(16, Cells(r, 19), " ", 1) + 1
'MsgBox quanti

Cells(r, 8).Copy Cells(r, 4)

Cells(r, 5) = ""

Cells(r, 6) = Mid(str, kao, lenkao)
Cells(r, 7) = ""
Cells(r, 8) = Mid(Cells(r, 19), xie, lenxie)

Cells(r, 9) = ""
Cells(r, 12) = Cells(r, 6)

Cells(r, 13) = Mid(Cells(r, 19), cargo, lencar)

Cells(r, 14) = Mid(Cells(r, 19), quanti, 6)

Cells(r, 19) = ""
Cells(r, c) = Left(str, 5)
endsub:
'航次更新结束

Cells(r, c).Select
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
