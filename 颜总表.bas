Attribute VB_Name = "颜总表"
Sub 名单分列预处理()

Set tou = ActiveCell
If tou = "" Then
    MsgBox "请选中名单第一名"
    Exit Sub
End If
If Left(tou.Text, InStr(1, tou.Text, ".") - 1) <> 1 Then
    MsgBox "请选中名单第一名"
    Exit Sub
Else
    If Cells(tou.Row + 1, tou.Column) = "" Then
        Set wei = tou
    Else
        Set wei = tou.End(xlDown)
    End If
End If
For i = 1 To wei.Row - tou.Row + 1
    rrr = Cells(tou.Row + i - 1, 1)
    If i < 10 Then
        If Mid(rrr, 2, 1) <> "." Then
            Cells(tou.Row + i - 1, 1) = Left(rrr, 1) & " " & Mid(rrr, 2, 555)
        End If
    Else
        If Mid(rrr, 3, 1) <> "." Then
            Cells(tou.Row + i - 1, 1) = Left(rrr, 2) & " " & Mid(rrr, 3, 555)
        End If
    End If
Next i
    Range(tou, wei).Select
    Selection.Replace what:="，", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=".", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="．", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="(", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="（", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="）", Replacement:=")", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="、", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="。", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="，", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="：", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=":", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="费", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
For i = tou.Row To wei.Row
Set rng1 = Cells(i, 1)
    strin = rng1.Text
    spcNum = InStr(2, rng1, " ")
    If InStr(4, rng1, "位") Then
        weiNum = InStr(4, rng1, "位")
        rng1.Value = Left(strin, spcNum) & Mid(strin, weiNum + 1, Len(strin) - weiNum)
    End If
Next i
Range(tou, wei).Select
        'A列按空格数据分列
    Selection.TextToColumns Destination:=tou, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
        '关掉数据分列
    Selection.TextToColumns Destination:=tou, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :="", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
End Sub '预处理，用来整理掉序号错乱
Sub 分解表格()
Call fenjie
End Sub
Sub 处理颜总的接龙报名()
'
' 处理颜总的接龙报名 宏
'

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Set tou = ActiveCell
If tou = 1 Then
    Set wei = tou.End(xlDown)
    GoTo chulihao
End If
If tou = "" Then
    MsgBox "请选中名单第一名"
    GoTo endsub:
End If
If Left(tou.Text, InStr(1, tou.Text, ".") - 1) <> 1 Then
    MsgBox "请选中名单第一名"
    GoTo endsub:
Else
    If Cells(tou.Row + 1, tou.Column) = "" Then
        Set wei = tou
    Else
        Set wei = tou.End(xlDown)
    End If
End If

For i = 1 To wei.Row - tou.Row + 1
    rrr = Cells(tou.Row + i - 1, 1)
    If i < 10 Then
        If Mid(rrr, 2, 1) <> "." Then
            Cells(tou.Row + i - 1, 1) = Left(rrr, 1) & " " & Mid(rrr, 2, 555)
        End If
    Else
        If Mid(rrr, 3, 1) <> "." Then
            Cells(tou.Row + i - 1, 1) = Left(rrr, 2) & " " & Mid(rrr, 3, 555)
        End If
    End If
Next i
    Range(tou, wei).Select
    Selection.Replace what:="，", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=".", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="．", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="(", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="（", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="）", Replacement:=")", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="、", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="。", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="，", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="：", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=":", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
   ' Selection.Replace What:="已", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="费", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
For i = tou.Row To wei.Row
Set rng1 = Cells(i, 1)
    strin = rng1.Text
    spcNum = InStr(2, rng1, " ")
    If InStr(4, rng1, "位") Then
        weiNum = InStr(4, rng1, "位")
        rng1.Value = Left(strin, spcNum) & Mid(strin, weiNum + 1, Len(strin) - weiNum)
    End If
Next i
Range(tou, wei).Select
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
        
chulihao:
tourow = tou.Row
weirow = wei.Value

For i = tourow To tourow + weirow - 1
    If i > tourow Then
        '名单第二名及以后的付费信息放到第一名的后面列里，然后在最后放回到相对应的名字后
        Cells(tourow, i - tourow + 3) = Cells(i, Cells(i, 2).End(xlToRight).Column).Text
    End If
    b = Cells(i + 1, 1).Value
    aa = Cells(i, 1).Value
    If b = "" Then
        GoTo shouwei:
    End If
    diff = b - aa
    crtRow = i
    If diff > 1 Then
    '插入空行，加入序列号，加入人名
        For cishu = 1 To b - aa - 1
            Rows(i + 1 & ":" & i + 1).Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(i + 1, 1).Value = Cells(i, 1).Value + 1
            Cells(i + 1, 2) = "=INDIRECT(ADDRESS(" & crtRow & "," & cishu + 2 & "))"
            i = i + 1
       Next cishu
    End If
    Next i
shouwei:

Set name1 = Cells(tourow, 2)
Set name2 = Cells(tou.End(xlDown).Row, 2)
    Range(name1, name2).Copy
    Range(name1, name2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
For i = tourow + 1 To tourow + tou.End(xlDown).Value - 1 '把上面的付款信息粘贴回来
    Cells(i, 3) = Cells(tourow, i - tourow + 3)
Next i

Set name1 = Cells(tourow, 3)
Set name2 = Cells(tou.End(xlDown).Row, 3)
Debug.Print name1
Debug.Print name2
Range(name1, name2).Select
        'c列按)数据分列
    Selection.TextToColumns Destination:=name1, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar _
        :=")", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
        '关掉数据分列
    Selection.TextToColumns Destination:=name1, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :="", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
        

Set name1 = Cells(tourow, 5)
Set name2 = Cells(tou.End(xlDown).Row, 100)
Range(name1, name2).ClearContents

endsub:
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub fenjie()
Set zong = ActiveSheet
yi = "一，小巧玲珑身材组美女有大福利啦！东方美旗袍丽人国际俱乐部招募一支小巧玲珑身材的旗袍走秀队员零基础学习旗袍走秀。"
mingziyi = "小个子"
er = "二，大长腿美眉组有大福利啦！东方美旗袍丽人国际俱乐部招募一支身高165以上旗袍走秀队员零基础学习旗袍走秀。"
mingzier = "大长腿"
san = "三，大龄姐姐组大福利啦，东方美旗袍丽人国际俱乐部招募一支65岁左右最美夕阳红旗袍走秀队员零基础学习旗袍走秀。"
mingzisan = "大姐姐"
si = "四，福美组特胖姐妹大福利啦，东方美旗袍丽人国际俱乐部招募一支穿特大尺码旗袍走秀队员零基础学习旗袍走秀。"
mingzisi = "特胖"
wu = "五，喜欢奥黛旗袍的姐妹有大福利啦，东方美旗袍丽人国际俱乐部招募一支穿奥黛旗袍的走秀队员零基础学习时装走秀（注意：奥黛组走时装步）"
mingziwu = "奥黛"
liu = "六，爱好伦巴舞的姐妹们有福利啦！东方美旗袍丽人国际俱乐部招募一支穿旗袍跳伦巴舞的姐妹队伍零基础学习伦巴舞。"
mingziliu = "伦巴"
qi = "七，旗袍拉丁伦巴舞11月11日10:30~12:00开班，指导老师：田网妹"
mingziqi = "拉丁七"
ba = "  八，东方美欲组建旗袍拉丁伦巴艺术表演团队。请有拉丁伦巴基础的姐妹们速度报名。350元10次课1件秋冬款玫红丝绒旗袍。"
mingziba = "拉丁表演八"
jiuyi = "一、表演班(每周六13:30~5:00，十次课，送自选旗袍一件)"
mingzijiuyi = "表演班"
jiuer = "二、中级班(周六15:00~16:30，指导老师:魏祥珍)"
mingzijiuer = "中级班"
xulie = Array(yi, er, san, si, wu, liu, qi, ba, jiuyi, jiuer)
mingzi = Array(mingziyi, mingzier, mingzisan, mingzisi, mingziwu, mingziliu, mingziqi, mingziba, mingzijiuyi, mingzijiuer)
i = 1
For r = 1 To 250
    If Cells(r, 1) = xulie(i) Then
        rwei = Cells(r, 1).End(xlDown).Row
        Sheets.Add After:=Sheets(Sheets.Count) '
        Sheets(Sheets.Count).Name = mingzi(i)
        
        zong.Select
        Range(Cells(r, 1), Cells(rwei, 4)).Copy
        Sheets(mingzi(i)).Select
        Range("a1").Activate
        ActiveSheet.Paste
        zong.Select
        i = i + 1
        If i = 10 Then
            Exit For
        End If
    End If
Next r
End Sub

Sub 处理颜总的接龙报名neo()
'
' 处理颜总的接龙报名neo 宏
'

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Set tou = ActiveCell
dizhi = tou.Address
'Set tou = Application.InputBox("1：:", "输入第一名所在单元格名称", dizhi, , , , , 8)
If tou = "" Then
    MsgBox "请选中名单第一名"
    GoTo endsub:
End If
If Left(tou.Text, InStr(1, tou.Text, ".") - 1) <> 1 Then
    MsgBox "请选中名单第一名"
    GoTo endsub:
Else
    If Cells(tou.Row + 1, tou.Column) = "" Then
        Set wei = tou
    Else
        Set wei = tou.End(xlDown)
    End If
End If


    Range(tou, wei).Select
    Selection.Replace what:="，", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=".", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="．", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="(", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="（", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="、", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="。", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="，", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="：", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=":", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="已", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="费", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="给", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
For i = tou.Row To wei.Row
Set rng1 = Cells(i, 1)
    strin = rng1.Text
    spcNum = InStr(2, rng1, " ")
    If InStr(4, rng1, "位") Then
        weiNum = InStr(4, rng1, "位")
        rng1.Value = Left(strin, spcNum) & Mid(strin, weiNum + 1, Len(strin) - weiNum)
    End If
Next i
Range(tou, wei).Select
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
tourow = tou.Row
weirow = wei.Value

For i = tourow To tourow + weirow - 1
    If i > tourow Then
        '名单第二名及以后的付费信息放到第一名的后面列里，然后在最后放回到相对应的名字后
        Cells(tourow, i - tourow + 3) = Cells(i, Cells(i, 2).End(xlToRight).Column).Text
    End If
    b = Cells(i + 1, 1).Value
    a = Cells(i, 1).Value
    If b = "" Then
        GoTo shouwei:
    End If
    diff = b - a
    crtRow = i
    If diff > 1 Then
    '插入空行，加入序列号，加入人名
        For cishu = 1 To b - a - 1
            Rows(i + 1 & ":" & i + 1).Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(i + 1, 1).Value = Cells(i, 1).Value + 1
           ' cells(i+1,2) =
            Cells(i + 1, 2) = "=INDIRECT(ADDRESS(" & crtRow & "," & cishu + 2 & "))"
            i = i + 1
       Next cishu
    End If
    Next i
shouwei:

Set name1 = Cells(tourow, 2)
Set name2 = Cells(tou.End(xlDown).Row, 2)
    Range(name1, name2).Copy
    Range(name1, name2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
For i = tourow + 1 To tourow + tou.End(xlDown).Value - 1
Cells(i, 3) = Cells(tourow, i - tourow + 3)
Next i
Set name1 = Cells(tourow, 4)
Set name2 = Cells(tou.End(xlDown).Row, 100)
Range(name1, name2).ClearContents
endsub:
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub a()
Range("b2") = "asdasd"
Sheets(2).Range("b5").Copy Range("a99999").End(xlUp)
Range("a99999").End(xlUp).Select
End Sub
Sub 综合分表()
renshu = 0
For i = 1 To 8
    For Each rng In Worksheets(i).Range("a2:a6")
        If rng = "" Then
            GoTo nexti
        End If
        If rng = "序号" Then
            tourow = rng.Row
            weirow = Worksheets(i).Cells(tourow, 2).End(xlDown).Row
            renshu = renshu + Worksheets(i).Cells(weirow, 1)
            Debug.Print Worksheets(i).Cells(weirow, 1)
            Exit For
        End If
    Next rng
    rowzb = Worksheets(9).Range("a99999").End(xlUp).Row
    If i = 1 Then
        Cells(rowzb, 1) = Worksheets(i).Name
    Else
        Cells(rowzb + 1, 1) = Worksheets(i).Name
    End If
    rowzb = Worksheets(9).Range("a99999").End(xlUp).Row
    Worksheets(i).Select
    Range(Cells(tourow, 1), Cells(weirow, 5)).Select
    Selection.Copy
    Worksheets(9).Select
    Cells(rowzb + 1, 1).Select
    ActiveSheet.Paste
    

nexti:
Next i
weirow = Range("a99999").End(xlUp).Row + 1
Cells(weirow, 1) = "总人数"
Cells(weirow, 2) = renshu
'Worksheets(9).Select
End Sub
Sub kuabiao()
Debug.Print Worksheets(9).Name
End Sub
