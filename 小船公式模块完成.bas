Attribute VB_Name = "С����ʽģ��"
Sub С���ܱ�һ���㶨()
Attribute С���ܱ�һ���㶨.VB_ProcData.VB_Invoke_Func = " \n14"
'
' С��һ��
'

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False
shipNamArr = Array("����1", "����10", "����15", "����16", "����17", "����18", "����2", "����3", "����32", "����5", "����7", "����9")
lujin = "F:\�����ĵ�\�¶Ȱ�ȫЧ�潱��1609-1708��\201710\"
zong = "����ҵ�����ƻ���201710��-С���ܱ�.xlsx"
For Each shipNam In shipNamArr

    For Each rng In Sheets("ʱ�����ͳ�Ʊ�").Range("a5:a16")
        If Len(shipNam) = 3 Then
            If Mid(rng.Value, 4, 1) = Chr(10) Then
                If Trim(Left(rng.Value, 3)) = shipNam Then
                    zongshiRow = rng.Row
                    GoTo ZhaoDaoZongRowHou:
                End If
            End If
        Else
            If Trim(Left(rng.Value, 4)) = shipNam Then
                zongshiRow = rng.Row
                GoTo ZhaoDaoZongRowHou:
            End If
        End If
    Next rng
ZhaoDaoZongRowHou:
    For Each rng In Sheets("ҵ�����ͳ�Ʊ�").Range("a2:a19")
        If rng.Value = shipNam Then
            zongguanrow = rng.Row
            GoTo zhaodaoguanrow
        End If
    Next rng
zhaodaoguanrow:
    For Each rng In Sheets("������Чͳ�Ʊ�").Range("a4:a19")
        If rng.Value = shipNam Then
            zongtongrow = rng.Row
            GoTo zentongrow
        End If
    Next rng
zentongrow:
    shipBk = shipNam & ".xlsx"
    Workbooks.Open Filename:=lujin & shipBk

    For Each rng In Workbooks(shipBk).Sheets("ʱ�����ͳ�Ʊ�").Range("a2:a19")
        If Len(shipNam) = 3 Then
            If Mid(rng.Value, 4, 1) = Chr(10) Then
                If Trim(Left(rng.Value, 3)) = shipNam Then
                    Workbooks(shipBk).Sheets("ʱ�����ͳ�Ʊ�").Rows(rng.Row).Copy Workbooks(zong).Sheets("ʱ�����ͳ�Ʊ�").Rows(zongshiRow)
                    GoTo qufuzhi:
                End If
            End If
        Else
            If Trim(Left(rng.Value, 4)) = shipNam Then
                Workbooks(shipBk).Sheets("ʱ�����ͳ�Ʊ�").Rows(rng.Row).Copy Workbooks(zong).Sheets("ʱ�����ͳ�Ʊ�").Rows(zongshiRow)
                GoTo qufuzhi:
            End If
        End If
    Next rng

qufuzhi:


    For Each rng In Workbooks(shipBk).Sheets("ҵ�����ͳ�Ʊ�").Range("a2:f2")
        If InStr(rng.Value, "�ຽ��Ӫ��") Then
            colduo = rng.Column
            GoTo col
        End If
    Next rng
col:
    For Each rng In Workbooks(shipBk).Sheets("ҵ�����ͳ�Ʊ�").Range("a2:a18")
        If rng.Value = shipNam Then
            If Workbooks(shipBk).Sheets("ҵ�����ͳ�Ʊ�").Cells(rng.Row, colduo) = "" Then
                Workbooks(zong).Sheets("ҵ�����ͳ�Ʊ�").Cells(zongguanrow, 2) = "ԭ���"
            Else
                Workbooks(shipBk).Sheets("ҵ�����ͳ�Ʊ�").Cells(rng.Row, colduo).Copy Workbooks(zong).Sheets("ҵ�����ͳ�Ʊ�").Cells(zongguanrow, 2)
                GoTo guanrow
            End If
        End If
    Next rng
guanrow:
    For Each rng In Workbooks(shipBk).Sheets("������Ч����").Range("a2:a180")
        i = rng.Row
        
            If rng.Text = shipNam Then
                rySize = rng.MergeArea.Rows.Count
                rend = i + rySize - 1
                For e = i To rend
                    If Workbooks(shipBk).Sheets("������Ч����").Cells(e, 5) = "" Then
                        rMaxe = e - 1
                        GoTo ii
                    End If
                Next e
ii:
                For ii = i To rend
                    If Workbooks(shipBk).Sheets("������Ч����").Cells(ii, 9) = "" Then
                        rMaxii = ii - 1
                        GoTo m
                    End If
                Next ii
m:
                For m = i To rend
                    If Workbooks(shipBk).Sheets("������Ч����").Cells(m, 13) = "" Then
                        rMaxm = m - 1
                        GoTo rma
                    End If
                Next m
rma:
                rMax = WorksheetFunction.Max(rMaxe, rMaxii, rMaxm)
                rMin = i
                GoTo over
            End If
    Next rng
over:
    For Each rng In Workbooks(zong).Sheets("������Ч����").Range("a2:a180")
        i = rng.Row
        If rng.Text = shipNam Then
            rzSize = rng.MergeArea.Rows.Count
            charu = rySize - rzSize
                If charu > 0 Then
                    For konghang = 1 To charu
                        Workbooks(zong).Sheets("������Ч����").Cells(i + 1, 2).EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
                    Next konghang
                End If
            Workbooks(shipBk).Sheets("������Ч����").Activate
            xiaoguo = Workbooks(shipBk).Sheets("������Ч����").Range(Cells(rMin, 2), Cells(rMax, 13))
            Workbooks(zong).Sheets("������Ч����").Activate
            Workbooks(zong).Sheets("������Ч����").Range(Cells(i, 2), Cells(i + rMax - rMin - 1, 13)) = xiaoguo
        GoTo hangcitongji
        End If
    Next rng
hangcitongji:
    For Each rng In Workbooks(shipBk).Sheets("������Чͳ�Ʊ�").Range("a3:a18")
        If rng.Value = shipNam Then
        
            If Workbooks(shipBk).Sheets("������Чͳ�Ʊ�").Cells(rng.Row, 2) = "" Then
                Workbooks(zong).Sheets("������Чͳ�Ʊ�").Cells(zongtongrow, 2) = "��"
            Else
                Workbooks(shipBk).Sheets("������Чͳ�Ʊ�").Cells(rng.Row, 2).Copy Workbooks(zong).Sheets("������Чͳ�Ʊ�").Cells(zongtongrow, 2)
            End If
            
            If Workbooks(shipBk).Sheets("������Чͳ�Ʊ�").Cells(rng.Row, 4) = "" Then
                Workbooks(zong).Sheets("������Чͳ�Ʊ�").Cells(zongtongrow, 4) = "��"
            Else
                Workbooks(shipBk).Sheets("������Чͳ�Ʊ�").Cells(rng.Row, 4).Copy Workbooks(zong).Sheets("������Чͳ�Ʊ�").Cells(zongtongrow, 4)
            End If
            
            If Workbooks(shipBk).Sheets("������Чͳ�Ʊ�").Cells(rng.Row, 6) = "" Then
                Workbooks(zong).Sheets("������Чͳ�Ʊ�").Cells(zongtongrow, 6) = "��"
            Else
                Workbooks(shipBk).Sheets("������Чͳ�Ʊ�").Cells(rng.Row, 6).Copy Workbooks(zong).Sheets("������Чͳ�Ʊ�").Cells(zongtongrow, 6)
            End If
            
            GoTo tongrow
        End If
    Next rng
tongrow:
Windows(shipNam & ".xlsx").Close
Next shipNam
'������
Windows(zong).Activate
Sheets("ʱ�����ͳ�Ʊ�").Select
zongshiRow = 5
Range("N" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]=""NO TKC"",1000,RC[-2]-RC[-1])"
Range("S" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]-RC[-5]<0,0,RC[-1]-RC[-4])"
Range("V" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]=0,0,RC[-2]/RC[-1])"
Range("X" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]=""NO TKC"",1000,RC[-2]-RC[-1])"
Range("AC" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]-RC[-5]<0,0,RC[-1]-RC[-4])"
Range("AF" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]=0,0,RC[-2]/RC[-1])"
Range("AH" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]=""NO TKC"",1000,RC[-2]-RC[-1])"
Range("AM" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]-RC[-5]<0,0,RC[-1]-RC[-4])"
Range("AP" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]=0,0,RC[-2]/RC[-1])"
Range("AR" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]=""NO TKC"",1000,RC[-2]-RC[-1])"
Range("AW" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]-RC[-5]<0,0,RC[-1]-RC[-4])"
Range("AZ" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]=0,0,RC[-2]/RC[-1])"
Range("BB" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]=""NO TKC"",1000,RC[-2]-RC[-1])"
Range("BH" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-41]+RC[-31]+RC[-21]+RC[-11]+RC[-1]"
Range("BG" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-1]-RC[-4]"
Range("CD" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-1]/RC[-2]"
Range("CF" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-2]-RC[-1]"
Range("CH" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-1]/RC[-6]"
Range("CJ" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-2]-RC[-1]"
Range("CL" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-1]/RC[-10]"
Range("CR" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-6]-RC[-1]"
Range("CT" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-1]/RC[-18]"
Range("CV" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-2]-RC[-1]"
Range("CX" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-1]/RC[-22]"
Range("CZ" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-2]-RC[-1]"
Range("DA" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-21]+RC[-17]+RC[-9]+RC[-5]+RC[-1]"
Range("DE" & CStr(zongshiRow)).FormulaR1C1 = "=SUM(RC[-49],RC[-4],RC[-2],RC[-1])"
Range("DG" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-2]*RC[-1]"

    Range("L5").AutoFill Destination:=Range("L5:L16"), Type:=xlFillDefault
    Range("N5").AutoFill Destination:=Range("N5:N16"), Type:=xlFillDefault
    Range("S5").AutoFill Destination:=Range("S5:S16"), Type:=xlFillDefault
    Range("V5").AutoFill Destination:=Range("V5:V16"), Type:=xlFillDefault
    Range("X5").AutoFill Destination:=Range("X5:X16"), Type:=xlFillDefault
    Range("AC5").AutoFill Destination:=Range("AC5:AC16"), Type:=xlFillDefault
    Range("AF5").AutoFill Destination:=Range("AF5:AF16"), Type:=xlFillDefault
    Range("AH5").AutoFill Destination:=Range("AH5:AH16"), Type:=xlFillDefault
    Range("AM5").AutoFill Destination:=Range("AM5:AM16"), Type:=xlFillDefault
    Range("AP5").AutoFill Destination:=Range("AP5:AP16"), Type:=xlFillDefault
    Range("AR5").AutoFill Destination:=Range("AR5:AR16"), Type:=xlFillDefault
    Range("AW5").AutoFill Destination:=Range("AW5:AW16"), Type:=xlFillDefault
    Range("AZ5").AutoFill Destination:=Range("AZ5:AZ16"), Type:=xlFillDefault
    Range("BB5").AutoFill Destination:=Range("BB5:BB16"), Type:=xlFillDefault
    Range("BG5").AutoFill Destination:=Range("BG5:BG16"), Type:=xlFillDefault
    Range("BH5").AutoFill Destination:=Range("BH5:BH16"), Type:=xlFillDefault
    Range("CD5").AutoFill Destination:=Range("CD5:CD16"), Type:=xlFillDefault
    Range("CF5").AutoFill Destination:=Range("CF5:CF16"), Type:=xlFillDefault
    Range("CH5").AutoFill Destination:=Range("CH5:CH16"), Type:=xlFillDefault
    Range("CJ5").AutoFill Destination:=Range("CJ5:CJ16"), Type:=xlFillDefault
    Range("CL5").AutoFill Destination:=Range("CL5:CL16"), Type:=xlFillDefault
    Range("CR5").AutoFill Destination:=Range("CR5:CR16"), Type:=xlFillDefault
    Range("CT5").AutoFill Destination:=Range("CT5:CT16"), Type:=xlFillDefault
    Range("CV5").AutoFill Destination:=Range("CV5:CV16"), Type:=xlFillDefault
    Range("CX5").AutoFill Destination:=Range("CX5:CX16"), Type:=xlFillDefault
    Range("CZ5").AutoFill Destination:=Range("CZ5:CZ16"), Type:=xlFillDefault
    Range("DA5").AutoFill Destination:=Range("DA5:DA16"), Type:=xlFillDefault
    Range("DE5").AutoFill Destination:=Range("DE5:DE16"), Type:=xlFillDefault
    Range("DG5").AutoFill Destination:=Range("DG5:DG16"), Type:=xlFillDefault
'���ͳһ��ʽ
    Rows("16").Copy
    Rows("5:16").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Cells(5, 1).Select

    Sheets("ҵ�����ͳ�Ʊ�").Select
    Rows("14").Copy
    Rows("3:14").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Cells(3, 1).Select

    Sheets("������Чͳ�Ʊ�").Select
    Rows("16").Copy
    Rows("5:16").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Cells(5, 1).Select

Sheets("ʱ�����ͳ�Ʊ�").Select

Application.ScreenUpdating = 1
Application.DisplayAlerts = 1
End Sub
Sub �����ر�()
Attribute �����ر�.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �����ر� ��
'

'"����1", "����10",
lujin = "D:\9�·��¶ȱ���\201709\"
shipNamArr = Array("����15", "����16", "����17", "����18", "����2", "����3", "����32", "����5", "����7", "����9")
zong = "����ҵ�����ƻ���201709��-С���ܱ�.xls"

For Each shipNam In shipNamArr
    shipBk = shipNam & ".xls"
    Workbooks.Open Filename:=lujin & shipBk
    Workbooks(shipBk).SaveAs Filename:=lujin & shipNam & ".xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    xlsx = shipNam & ".xlsx"
    Workbooks(xlsx).Close
    
Next shipNam

End Sub
