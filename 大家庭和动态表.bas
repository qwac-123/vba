Attribute VB_Name = "���ͥ�Ͷ�̬��"
Sub ��ʼ��������Ĺ���()
'
'�򿪶�̬��ʹ��ͥ����ʾ��̬��

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False
If Weekday(Date, vbMonday) = 1 Then '�����������һ���ʹ�����ģ�����ʹ������
    zuotian = 3
Else
    zuotian = 1
End If
    Workbooks.Open Filename:=Format(Date - zuotian, "F:\\�����ĵ�\\��̬��������ͥ��\\mm��\\��̬��������ͥ��yyyy-mm-dd.xl\sx")
    Workbooks(Format(Date - zuotian, "��̬��������ͥ��yyyy-mm-dd.xl\sx")).Activate
        With Application '��������1��
        .Iteration = True
        .MaxIterations = 1
    End With
    
    Workbooks.Open Filename:= _
        "\\192.168.0.223\\��������\\3.2��������\\4 ������̬��\\" & Format(Date - zuotian, "yyyy\\mm��\\������̬��yyyy-mm-dd��.xl\sx")
Application.ScreenUpdating = 1
Application.DisplayAlerts = 1
Workbooks(Format(Date - zuotian, "������̬��yyyy-mm-dd��.xl\sx")).Activate
End Sub
Sub ���ͥ��̬()
'
' ��������̬ Macro
' ��������̬��Ϣ����K1����β��ӣ�Ȼ������J�в����У����/����ȥ����rob����
'

'
If Range("k1") = "" Then
MsgBox "k1���ǿյģ��ǲ������Ѿ����һ����"
Exit Sub
End If


Application.ScreenUpdating = False
Application.DisplayAlerts = False

   ' With Application '��������1��
  '      .Iteration = True
  '      .MaxIterations = 1
  '  End With
Range("d4:e16").Interior.Pattern = xlNone
    With Range("k1:k25") '�����ƹ����Ĵ�����̬
    .Replace What:="��", Replacement:=":", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="����", Replacement:="DH", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="����", Replacement:="JX", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="��", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="��", Replacement:="��", LookAt:=xlPart, _
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
endchulidongtai:    '��̬�������
Range("k1:k35").ClearContents
Range("j4:j15").Copy Range("k4:k15")
Range("a4:a15").Copy Range("j4:j15")
Range("a1") = "�Ϻ����⴬�Ӷ�̬��Ϣһ���� " & Format(Date, "yyyy��m��d�� aaaa") '�������� Range("G1:I1").FormulaR1C1 = "=IF(RC=0,TEXT(NOW(),""yyyy��m��d�� aaaa""),RC)"
    Range("F4:F15").FormulaR1C1 = _
        "=IF(RC[1]<>"""",""����""&MID(RC[1],5,3),IF(RC[2]<>"""",""ê��""&MID(RC[2],5,3),IF(COUNT(FIND(""����"",RC[5])),IF(SUM(ISNUMBER(FIND({""�żҸ�"",""���Ƹ�"",""����Ȧ"",""���˵�""},RC[5]))*1),MID(RC[5],FIND(""����"",RC[5]),5),MID(RC[5],FIND(""����"",RC[5]),4)),RC[6]&""���"")))"
    Range("h4:h15") = ""
    
    

Set djt = ActiveWorkbook
Range("K4:k15").Copy '���ͥ������̬���Ƶ���̬��
Call check_wk������̬

    Range("J4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("L4").Select
        
    
djt.Activate
Range("h4").Select




Application.ScreenUpdating = True
Application.DisplayAlerts = True
    ActiveWorkbook.SaveAs Filename:= _
        Format(Date, "F:\\�����ĵ�\\��̬��������ͥ��\\mm��\\��̬��������ͥ��yyyy-mm-dd.xl\sx"), FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
End Sub

Sub check_wk������̬()
On Error Resume Next
If Weekday(Date, vbMonday) = 1 Then '�����������һ���ʹ�����ģ�����ʹ������
    zuotian = 3
Else
    zuotian = 1
End If
Set wkjin = Workbooks(Format(Date, "������̬��yyyy-mm-dd��.xl\sx"))
Set wkzuo = Workbooks(Format(Date - zuotian, "������̬��yyyy-mm-dd��.xl\sx"))
If Len(wkjin.Name) > 0 Then
    If Err.Number = 9 Then '���wkjin���ˣ��ͼ���
        wkjin.Activate
    Else                    '����ͼ�����һ��
        wkzuo.Activate
    End If
End If
End Sub

Sub bbb��̬��ISMSROB()
'
' ����ISMSROB
'
'
Call check_wk������̬
Worksheets("Vessel Status").Activate
Dim i As Integer
Application.ScreenUpdating = False
Application.DisplayAlerts = False
'����ISMSrob

    Range("o4:o15").Select
    Selection.Replace What:="��", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="ˮ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��", Replacement:="MT", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    'Selection.Replace What:="��", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    'Selection.Replace What:="��", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    'Selection.Replace What:="��", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


For i = 4 To 15 Step 1
    ROB = Cells(i, 15).Text 'column "O"
    If InStr(1, ROB, ":") > 0 Then
        Cells(i, 15) = Mid(ROB, 9, InStr(11, ROB, "t") - 8) 'fo
        'Debug.Print Mid(rob, 9, InStr(11, rob, "t") - 8)
      Cells(i, 16) = Mid(ROB, InStr(16, ROB, ":") + 1, InStr(19, ROB, "t") - InStr(16, ROB, ":")) 'do
      'Debug.Print Mid(rob, InStr(16, rob, ":") + 1, InStr(19, rob, "t") - InStr(16, rob, ":")) 'do
      Cells(i, 17) = Mid(ROB, InStr(40, ROB, ":") + 1, InStr(46, ROB, "L") - InStr(40, ROB, ":")) 'lo
      'Debug.Print Mid(rob, InStr(40, rob, ":") + 1, InStr(46, rob, "L") - InStr(40, rob, ":")) 'lo
      Cells(i, 18) = Mid(ROB, InStr(27, ROB, ":") + 1, InStr(30, ROB, "t") - InStr(27, ROB, ":")) 'fw
      'Debug.Print Mid(rob, InStr(27, rob, ":") + 1, InStr(30, rob, "t") - InStr(27, rob, ":")) 'fw
    End If
    
    If InStr(1, ROB, "/") > 0 Then
        '����rob
        Cells(i, 15).TextToColumns Destination:=Cells(i, 15), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    
        '�رշ���
        Range("o4").TextToColumns Destination:=Range("o4"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :="", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=0
    End If
Next i
'�ָ�rob��ʽ
    Range("O16").Copy
    Range("O4:R16").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    

    Sheets("agent info.").Range("a1").FormulaR1C1 = "���ӵ�ǰ���δ�����Ϣ( " & Format(Date, "yyyy��m��d�� aaaa") & ")" '��������"=IF(RC=0,TEXT(NOW(),""yyyy��m��d�� aaaa""&"")""),RC)"
    Sheets("coordinate info.").Range("a1").FormulaR1C1 = "ҵ��Э����������(" & Format(Date, "yyyy��m��d�� aaaa") & ")"  '�������� "=IF(RC=0,TEXT(NOW(),""yyyy��m��d�� aaaa""&"")""),RC)"
'��ʼ����γ��

Sub Macro1()
'
' Macro1 Macro
'

'
    Windows("PERSONAL.xlsb").Activate
    Sheets("��λ����").Select
If Not Range("k5") = "" Then

    Range("K21:K35").Copy
    Windows("������̬��2017-11-09��.xlsx").Activate
    Range("K10").Select
    Range("K4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K4").Select
Else
  Windows("������̬��2017-11-09��.xlsx").Activate
MsgBox "��û���ƴ�����γ��"
End If
End Sub


Application.ScreenUpdating = True
Application.DisplayAlerts = True
'�������

    ActiveWorkbook.SaveAs Filename:= _
        "\\192.168.0.223\\��������\\3.2��������\\4 ������̬��\\" & Format(Date, "yyyy\\mm��\\������̬��yyyy-mm-dd��.xl\sx"), _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub
Sub ���ͥ�º���()
r = ActiveCell.Row
Set loadP = Cells(r, 4)
Set discP = Cells(r, 5)
Set NextV = Cells(r, 9)
Cells(r, 4) = Left(NextV.Value, InStr(3, NextV.Value, "-") - 1)
Cells(r, 5) = Right(NextV.Value, Len(NextV.Value) - InStr(3, NextV.Value, "-"))
NextV = "����"
End Sub
Sub aaa�����º���()
Dim kaishi, jieshu, i As Integer, str, abc As String

r = ActiveCell.Row
c = ActiveCell.Column
Application.ScreenUpdating = 0
Application.DisplayAlerts = 0
'����ͳһ���������ʽ
    Cells(r, 2).Select
    Selection.Replace What:="v", Replacement:="V", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells(r, 19).Select
    Selection.Replace What:="v", Replacement:="V", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=Chr(10), Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��", Replacement:="(", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��", Replacement:=")", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="--", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=",", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��", Replacement:="MT", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��", Replacement:="(", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��", Replacement:=")", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
  '�ʺŻ��滻��������  'Selection.Replace What:="?", Replacement:=" ", LookAt:=xlPart, _
    '    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
      '  ReplaceFormat:=False
    Selection.Replace What:=" ", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="��5%", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

str = Trim(Cells(r, 19).Text)
If str = "" Then
Exit Sub
End If

If Left(str, 1) <> "V" Then
str = "V" & str
'MsgBox "�жϲ���ӿ�ͷV��" & str
End If

If Mid(str, 6, 1) = "&" Then
xiahangci = Right(str, Len(str) - 6)
str = Left(str, 5) & Right(str, Len(str) - 10) '���V1234&�У��Ͱ�&����Ų��
End If
If Len(str) > 49 Then
xiahangci = Right(str, Len(str) - InStr(17, str, "V1") + 1)
str = Trim(Left(str, InStr(17, str, "V1") - 1)) '���̫����˵��������������Ϣ������ĺ�����ϢŲ��
End If
If Mid(str, 6, 1) = " " Then
str = Left(str, 5) & Right(str, Len(str) - 6) 'ȥ��V1234 ������Ŀո�
End If
If Mid(str, 6, 1) <> "��" Then
str = Left(str, 5) & "��������" & Right(str, Len(str) - 5) 'v1234װж������Ϻ�������
End If
If Mid(str, 10, 1) = " " Then
str = Left(str, 9) & Right(str, Len(str) - 10) 'v1234�������� װж ȥ���ո�
End If
If Mid(str, 10, 1) <> "(" Then
str = Left(str, 9) & "(" & Right(str, Len(str) - 9) 'MsgBox "����(��" & str
End If
If InStr(9, Left(str, Len(str) - 2), ")") <> 0 Then
str = Left(str, InStr(9, str, ")") - 1) & Right(str, Len(str) - InStr(9, str, ")")) 'MsgBox "ɾ������ǰ)��" & str
End If
If Mid(str, InStr(13, str, "T") - 6, 1) <> " " Then
str = Left(str, InStr(13, str, "T") - 6) & " " & Right(str, Len(str) - InStr(13, str, "T") + 6)
'MsgBox "����ǰ�ӿո�" & str
End If
If Mid(str, InStr(13, str, "T") + 1, 1) <> " " Then
str = Left(str, InStr(13, str, "T")) & " " & Right(str, Len(str) - InStr(13, str, "T"))
'MsgBox "����ǰ�ӿո�" & str
End If

If Right(str, 1) = "��" Then
str = Left(str, Len(str) - 4)
'MsgBox "ɾ�����ĺ������" & str
End If
If Right(str, 1) <> ")" Then
str = str & ")"
'MsgBox "����)��" & str
End If
Cells(r, 19) = str
        
 '�������
 
' MsgBox "�������" & i
 
 '���ν��������´������мƻ�

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
'���θ��½���

Cells(r, c).Select
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
