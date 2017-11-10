Attribute VB_Name = "大家庭和动态表"
Sub 开始今天上午的工作()
'
'打开动态表和大家庭并显示动态表

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False
If Weekday(Date, vbMonday) = 1 Then '如果今天是周一，就打开周五的；否则就打开昨天的
    zuotian = 3
Else
    zuotian = 1
End If
    Workbooks.Open Filename:=Format(Date - zuotian, "F:\\工作文档\\动态表（鼎衡大家庭）\\mm月\\动态表（鼎衡大家庭）yyyy-mm-dd.xl\sx")
    Workbooks(Format(Date - zuotian, "动态表（鼎衡大家庭）yyyy-mm-dd.xl\sx")).Activate
        With Application '开启迭代1次
        .Iteration = True
        .MaxIterations = 1
    End With
    
    Workbooks.Open Filename:= _
        "\\192.168.0.223\\航运在线\\3.2、操作部\\4 船舶动态表\\" & Format(Date - zuotian, "yyyy\\mm月\\船舶动态（yyyy-mm-dd）.xl\sx")
Application.ScreenUpdating = 1
Application.DisplayAlerts = 1
Workbooks(Format(Date - zuotian, "船舶动态（yyyy-mm-dd）.xl\sx")).Activate
End Sub
Sub zhoumo()
Debug.Print Weekday(Date, vbMonday)
End Sub
Sub 大家庭动态()
'
' 处理船舶动态 Macro
' 将船舶动态信息贴在K1格，首尾相接，然后整理到J列并分列，后接/分列去处理rob数据
'

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False

   ' With Application '开启迭代1次
  '      .Iteration = True
  '      .MaxIterations = 1
  '  End With
Range("d4:e16").Interior.Pattern = xlNone
    With Range("k1:k25") '处理复制过来的船舶动态
    .Replace What:="：", Replacement:=":", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="鼎衡", Replacement:="DH", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="轮", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="妥", Replacement:="泊", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    End With


For Each rngk In Range("k1:k25")
    j = 4
    i = 0
    If InStr(1, rngk, ":") = 0 Then
        If i > 2 Then
            GoTo endchulidongtai
        End If
        i = i + 1
        GoTo nextrngk
    End If
    xinxi = Mid(rngk.Text, InStr(1, rngk.Text, ":") + 1, 999)
    xinxitou = Mid(rngk.Text, 1, InStr(4, rngk.Text, ":") - 1)
    For Each rnga In Range("a4:a20")
        If xinxitou = rnga Then
            Cells(j, 10) = xinxi
            GoTo nextrngk
        End If
        j = j + 1
    Next rnga
nextrngk:
Next rngk
endchulidongtai:    '动态处理完毕
Range("k1:k35").ClearContents
Range("j4:j15").Copy Range("k4:k15")
Range("a4:a15").Copy Range("j4:j15")
    Range("G1:I1").FormulaR1C1 = "=IF(RC=0,TEXT(NOW(),""yyyy年m月d日 aaaa""),RC)" '更新日期

    Range("F4:F15").FormulaR1C1 = _
        "=IF(RC[1]<>"""",""开往""&MID(RC[1],5,3),IF(RC[2]<>"""",""锚泊""&MID(RC[2],5,3),IF(COUNT(FIND(""靠泊"",RC[5])),IF(SUM(ISNUMBER(FIND({""张家港"",""连云港"",""鲅鱼圈"",""仙人岛""},RC[5]))*1),MID(RC[5],FIND(""靠泊"",RC[5]),5),MID(RC[5],FIND(""靠泊"",RC[5]),4)),RC[6]&""完货"")))"
    Range("h4:h15") = ""
    
    
    
Range("h4").Select


Application.ScreenUpdating = True
Application.DisplayAlerts = True
    ActiveWorkbook.SaveAs Filename:= _
        Format(Date, "F:\\工作文档\\动态表（鼎衡大家庭）\\mm月\\动态表（鼎衡大家庭）yyyy-mm-dd.xl\sx"), FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
Call 动态表ISMSROB
End Sub

Sub 动态表ISMSROB()
'
' 处理ISMSROB
'
'
Workbooks(Format(Date - 1, "船舶动态（yyyy-mm-dd）.xl\sx")).Activate

Dim i As Integer
Application.ScreenUpdating = False
Application.DisplayAlerts = False
'处理ISMSrob
For i = 4 To 13 Step 1
    rob = Cells(i, 15).Text 'column "O"
    If Len(rob) < 50 Then
        GoTo nxt
    End If
    Cells(i, 15) = Mid(rob, 9, InStr(13, rob, "t") - 8) 'fo
    'Debug.Print Mid(rob, 9, InStr(13, rob, "t") - 8)
    Cells(i, 16) = Mid(rob, InStr(16, rob, ":") + 1, InStr(19, rob, "t") - InStr(16, rob, ":")) 'do
    'Debug.Print Mid(rob, InStr(16, rob, ":") + 1, InStr(19, rob, "t") - InStr(16, rob, ":")) 'do
    Cells(i, 17) = Mid(rob, InStr(40, rob, ":") + 1, InStr(46, rob, "L") - InStr(40, rob, ":")) 'lo
    'Debug.Print Mid(rob, InStr(40, rob, ":") + 1, InStr(46, rob, "L") - InStr(40, rob, ":")) 'lo
    Cells(i, 18) = Mid(rob, InStr(27, rob, ":") + 1, InStr(30, rob, "t") - InStr(27, rob, ":")) 'fw
    'Debug.Print Mid(rob, InStr(27, rob, ":") + 1, InStr(30, rob, "t") - InStr(27, rob, ":")) 'fw
nxt:
    Next

'只剩下dh7,jx32没有ISMS
For i = 14 To 15 Step 1
If Len(Range("o" & i)) > 15 Then
    '分列rob
    Range("o" & i).TextToColumns Destination:=Range("o" & i), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
End If
Next i
'关闭分列
    Range("o4").TextToColumns Destination:=Range("o4"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :="", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=0
'恢复rob格式
    Range("O16").Copy
    Range("O4:R16").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    

    Sheets("agent info.").Range("H1:I1").FormulaR1C1 = "=IF(RC=0,TEXT(NOW(),""yyyy年m月d日 aaaa""&"")""),RC)"
    Sheets("coordinate info.").Range("D1").FormulaR1C1 = "=IF(RC=0,TEXT(NOW(),""yyyy年m月d日 aaaa""&"")""),RC)"
'开始处理经纬度
    Sheets("Vessel Status").Range("L35").Select
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'保存今天

    ActiveWorkbook.SaveAs Filename:= _
        "\\192.168.0.223\\航运在线\\3.2、操作部\\4 船舶动态表\\" & Format(Date, "yyyy\\mm月\\船舶动态（yyyy-mm-dd）.xl\sx"), _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub

Sub 大家庭下航次()
r = ActiveCell.Row
Set loadP = Cells(r, 4)
Set discP = Cells(r, 5)
Set NextV = Cells(r, 9)
Cells(r, 4) = Left(NextV.Value, InStr(3, NextV.Value, "-") - 1)
Cells(r, 5) = Right(NextV.Value, Len(NextV.Value) - InStr(3, NextV.Value, "-"))
NextV = "待告"
End Sub

Sub 快速下航次()
Dim kaishi, jieshu, i As Integer, str, abc As String

r = ActiveCell.Row
c = ActiveCell.Column

Application.ScreenUpdating = 0
Application.DisplayAlerts = 0
'首先统一航次命令格式
    Range("b4:b16").Replace What:="v", Replacement:="V", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("s4:s16").Select
    Selection.Replace What:="v", Replacement:="V", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=Chr(10), Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="【", Replacement:="(", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="】", Replacement:=")", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="—", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="--", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="，", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=",", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="吨", Replacement:="MT", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="（", Replacement:="(", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="）", Replacement:=")", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="±5%", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


For r = 4 To 15
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
Next r
'航次更新结束

Cells(r, c).Select
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub


Sub test1()
  Dim response
  response = MsgBox("请点击下面的按钮", vbOKCancel)
  If response = 1 Then
    MsgBox "你点击了确定，很棒" '如果response的值是1，执行这条语句
  Else
    Exit Sub '如果response的值不是1，退出过程
  End If
End Sub
Sub test2()
Debug.Print res = MsgBox("弹窗演示", vbInformation + vbOKOnly, "这是一个弹出")
End Sub

