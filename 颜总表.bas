Attribute VB_Name = "���ܱ�"
Sub ��������Ԥ����()

Set tou = ActiveCell
If tou = "" Then
    MsgBox "��ѡ��������һ��"
    Exit Sub
End If
If Left(tou.Text, InStr(1, tou.Text, ".") - 1) <> 1 Then
    MsgBox "��ѡ��������һ��"
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
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=".", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="(", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=")", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=":", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
For i = tou.Row To wei.Row
Set rng1 = Cells(i, 1)
    strin = rng1.Text
    spcNum = InStr(2, rng1, " ")
    If InStr(4, rng1, "λ") Then
        weiNum = InStr(4, rng1, "λ")
        rng1.Value = Left(strin, spcNum) & Mid(strin, weiNum + 1, Len(strin) - weiNum)
    End If
Next i
Range(tou, wei).Select
        'A�а��ո����ݷ���
    Selection.TextToColumns Destination:=tou, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
        '�ص����ݷ���
    Selection.TextToColumns Destination:=tou, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :="", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
End Sub 'Ԥ���������������Ŵ���
Sub �ֽ���()
Call fenjie
End Sub
Sub �������ܵĽ�������()
'
' �������ܵĽ������� ��
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
    MsgBox "��ѡ��������һ��"
    GoTo endsub:
End If
If Left(tou.Text, InStr(1, tou.Text, ".") - 1) <> 1 Then
    MsgBox "��ѡ��������һ��"
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
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=".", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="(", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=")", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=":", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
   ' Selection.Replace What:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
For i = tou.Row To wei.Row
Set rng1 = Cells(i, 1)
    strin = rng1.Text
    spcNum = InStr(2, rng1, " ")
    If InStr(4, rng1, "λ") Then
        weiNum = InStr(4, rng1, "λ")
        rng1.Value = Left(strin, spcNum) & Mid(strin, weiNum + 1, Len(strin) - weiNum)
    End If
Next i
Range(tou, wei).Select
        'A�а��ո����ݷ���
    Selection.TextToColumns Destination:=tou, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="(", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
        '�ص����ݷ���
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
        '�����ڶ������Ժ�ĸ�����Ϣ�ŵ���һ���ĺ������Ȼ�������Żص����Ӧ�����ֺ�
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
    '������У��������кţ���������
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
    
For i = tourow + 1 To tourow + tou.End(xlDown).Value - 1 '������ĸ�����Ϣճ������
    Cells(i, 3) = Cells(tourow, i - tourow + 3)
Next i

Set name1 = Cells(tourow, 3)
Set name2 = Cells(tou.End(xlDown).Row, 3)
Debug.Print name1
Debug.Print name2
Range(name1, name2).Select
        'c�а�)���ݷ���
    Selection.TextToColumns Destination:=name1, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar _
        :=")", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
        '�ص����ݷ���
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
yi = "һ��С�������������Ů�д��������������������˹��ʾ��ֲ���ļһ֧С��������ĵ����������Ա�����ѧϰ�������㡣"
mingziyi = "С����"
er = "����������ü���д��������������������˹��ʾ��ֲ���ļһ֧���165�������������Ա�����ѧϰ�������㡣"
mingzier = "����"
san = "��������������������������������˹��ʾ��ֲ���ļһ֧65����������Ϧ�������������Ա�����ѧϰ�������㡣"
mingzisan = "����"
si = "�ģ����������ֽ��ô��������������������˹��ʾ��ֲ���ļһ֧���ش�������������Ա�����ѧϰ�������㡣"
mingzisi = "����"
wu = "�壬ϲ���������۵Ľ����д��������������������˹��ʾ��ֲ���ļһ֧���������۵������Ա�����ѧϰʱװ���㣨ע�⣺��������ʱװ����"
mingziwu = "����"
liu = "���������װ���Ľ������и��������������������˹��ʾ��ֲ���ļһ֧���������װ���Ľ��ö��������ѧϰ�װ��衣"
mingziliu = "�װ�"
qi = "�ߣ����������װ���11��11��10:30~12:00���ָ࣬����ʦ��������"
mingziqi = "������"
ba = "  �ˣ����������齨���������װ����������Ŷӡ����������װͻ����Ľ������ٶȱ�����350Ԫ10�ο�1���ﶬ��õ��˿�����ۡ�"
mingziba = "�������ݰ�"
jiuyi = "һ�����ݰ�(ÿ����13:30~5:00��ʮ�οΣ�����ѡ����һ��)"
mingzijiuyi = "���ݰ�"
jiuer = "�����м���(����15:00~16:30��ָ����ʦ:κ����)"
mingzijiuer = "�м���"
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

Sub �������ܵĽ�������neo()
'
' �������ܵĽ�������neo ��
'

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Set tou = ActiveCell
dizhi = tou.Address
'Set tou = Application.InputBox("1��:", "�����һ�����ڵ�Ԫ������", dizhi, , , , , 8)
If tou = "" Then
    MsgBox "��ѡ��������һ��"
    GoTo endsub:
End If
If Left(tou.Text, InStr(1, tou.Text, ".") - 1) <> 1 Then
    MsgBox "��ѡ��������һ��"
    GoTo endsub:
Else
    If Cells(tou.Row + 1, tou.Column) = "" Then
        Set wei = tou
    Else
        Set wei = tou.End(xlDown)
    End If
End If


    Range(tou, wei).Select
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=".", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="(", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:=":", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace what:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
For i = tou.Row To wei.Row
Set rng1 = Cells(i, 1)
    strin = rng1.Text
    spcNum = InStr(2, rng1, " ")
    If InStr(4, rng1, "λ") Then
        weiNum = InStr(4, rng1, "λ")
        rng1.Value = Left(strin, spcNum) & Mid(strin, weiNum + 1, Len(strin) - weiNum)
    End If
Next i
Range(tou, wei).Select
        'A�а��ո����ݷ���
    Selection.TextToColumns Destination:=tou, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="(", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
        '�ص����ݷ���
    Selection.TextToColumns Destination:=tou, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :="", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
tourow = tou.Row
weirow = wei.Value

For i = tourow To tourow + weirow - 1
    If i > tourow Then
        '�����ڶ������Ժ�ĸ�����Ϣ�ŵ���һ���ĺ������Ȼ�������Żص����Ӧ�����ֺ�
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
    '������У��������кţ���������
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
Sub �ۺϷֱ�()
renshu = 0
For i = 1 To 8
    For Each rng In Worksheets(i).Range("a2:a6")
        If rng = "" Then
            GoTo nexti
        End If
        If rng = "���" Then
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
Cells(weirow, 1) = "������"
Cells(weirow, 2) = renshu
'Worksheets(9).Select
End Sub
Sub kuabiao()
Debug.Print Worksheets(9).Name
End Sub
