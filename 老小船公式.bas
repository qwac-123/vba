Attribute VB_Name = "小船公式模块"

Sub 小船总表一键搞定()
Attribute 小船总表一键搞定.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 小船一键
'

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False
lujin = "D:\9月份月度报表\201709\"
shipNamArr = Array("鼎衡1", "鼎衡10", "鼎衡15", "鼎衡16", "鼎衡17", "鼎衡18", "鼎衡2", "鼎衡3", "建兴32", "鼎衡5", "鼎衡7", "鼎衡9")
zong = "船队业务管理计划【201709】-小船总表.xls"
For Each shipNam In shipNamArr

    For Each Rng In Sheets("时间管理统计表").Range("a5:a16")
        If Len(shipNam) = 3 Then
            If Mid(Rng.Value, 4, 1) = Chr(10) Then
                If Trim(Left(Rng.Value, 3)) = shipNam Then
                    zongshiRow = Rng.Row
                    GoTo ZhaoDaoZongRowHou:
                End If
            End If
        Else
            If Trim(Left(Rng.Value, 4)) = shipNam Then
                zongshiRow = Rng.Row
                GoTo ZhaoDaoZongRowHou:
            End If
        End If
    Next Rng
ZhaoDaoZongRowHou:
    For Each Rng In Sheets("业务管理统计表").Range("a3:a14")
        If Rng.Value = shipNam Then
            zongguanrow = Rng.Row
            GoTo zhaodaoguanrow
        End If
    Next Rng
zhaodaoguanrow:
    For Each Rng In Sheets("航次增效统计表").Range("a4:a17")
        If Rng.Value = shipNam Then
            zongtongrow = Rng.Row
            GoTo zentongrow
        End If
    Next Rng
zentongrow:
Debug.Print zongshiRow
Debug.Print zongguanrow
Debug.Print zongtongrow
    shipBk = shipNam & ".xlsx"
    Workbooks.Open Filename:=lujin & shipBk

    For Each Rng In Workbooks(shipBk).Sheets("时间管理统计表").Range("a3:a17")
        If Len(shipNam) = 3 Then
            If Mid(Rng.Value, 4, 1) = Chr(10) Then
                If Trim(Left(Rng.Value, 3)) = shipNam Then
                    Workbooks(shipBk).Sheets("时间管理统计表").Rows(Rng.Row).Copy Workbooks(zong).Sheets("时间管理统计表").Rows(zongshiRow)
                    GoTo qufuzhi:
                End If
            End If
        Else
            If Trim(Left(Rng.Value, 4)) = shipNam Then
                Workbooks(shipBk).Sheets("时间管理统计表").Rows(Rng.Row).Copy Workbooks(zong).Sheets("时间管理统计表").Rows(zongshiRow)
                GoTo qufuzhi:
            End If
        End If
    Next Rng

qufuzhi:


    For Each Rng In Workbooks(shipBk).Sheets("业务管理统计表").Range("a2:f2")
        If InStr(Rng.Value, "多航次营运") Then
            colduo = Rng.Column
            GoTo col
        End If
    Next Rng
col:
    For Each Rng In Workbooks(shipBk).Sheets("业务管理统计表").Range("a3:a14")
        If Rng.Value = shipNam Then
            If Workbooks(shipBk).Sheets("业务管理统计表").Cells(Rng.Row, colduo) = "" Then
                Workbooks(zong).Sheets("业务管理统计表").Cells(zongguanrow, 2) = "原表空"
            Else
                Workbooks(shipBk).Sheets("业务管理统计表").Cells(Rng.Row, colduo).Copy Workbooks(zong).Sheets("业务管理统计表").Cells(zongguanrow, 2)
                GoTo guanrow
            End If
        End If
    Next Rng
guanrow:

    For Each Rng In Workbooks(shipBk).Sheets("航次增效统计表").Range("a3:a14")
        If Rng.Value = shipNam Then
        
            If Workbooks(shipBk).Sheets("航次增效统计表").Cells(Rng.Row, 2) = "" Then
                Workbooks(zong).Sheets("航次增效统计表").Cells(zongtongrow, 2) = "空"
            Else
                Workbooks(shipBk).Sheets("航次增效统计表").Cells(Rng.Row, 2).Copy Workbooks(zong).Sheets("航次增效统计表").Cells(zongtongrow, 2)
            End If
            
            If Workbooks(shipBk).Sheets("航次增效统计表").Cells(Rng.Row, 4) = "" Then
                Workbooks(zong).Sheets("航次增效统计表").Cells(zongtongrow, 4) = "空"
            Else
                Workbooks(shipBk).Sheets("航次增效统计表").Cells(Rng.Row, 4).Copy Workbooks(zong).Sheets("航次增效统计表").Cells(zongtongrow, 4)
            End If
            
            If Workbooks(shipBk).Sheets("航次增效统计表").Cells(Rng.Row, 6) = "" Then
                Workbooks(zong).Sheets("航次增效统计表").Cells(zongtongrow, 6) = "空"
            Else
                Workbooks(shipBk).Sheets("航次增效统计表").Cells(Rng.Row, 6).Copy Workbooks(zong).Sheets("航次增效统计表").Cells(zongtongrow, 6)
            End If
            
            GoTo tongrow
        End If
tongrow:
    Next Rng

Next shipNam



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
Application.ScreenUpdating = 1
Application.DisplayAlerts = 1
End Sub
Sub 打开另存关闭()
Attribute 打开另存关闭.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 打开另存关闭 宏
'

'"鼎衡1", "鼎衡10",
lujin = "D:\9月份月度报表\201709\"
shipNamArr = Array("鼎衡15", "鼎衡16", "鼎衡17", "鼎衡18", "鼎衡2", "鼎衡3", "建兴32", "鼎衡5", "鼎衡7", "鼎衡9")
zong = "船队业务管理计划【201709】-小船总表.xls"

For Each shipNam In shipNamArr
    shipBk = shipNam & ".xls"
    Workbooks.Open Filename:=lujin & shipBk
    Workbooks(shipBk).SaveAs Filename:=lujin & shipNam & ".xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    xlsx = shipNam & ".xlsx"
    Workbooks(xlsx).Close
    
Next shipNam

End Sub
 
Sub 宏10()
Attribute 宏10.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏10 宏
'

'\
shipNamArr = Array("鼎衡15", "鼎衡16", "鼎衡17", "鼎衡18", "鼎衡2", "鼎衡3", "建兴32", "鼎衡5", "鼎衡7", "鼎衡9")
For Each nam In shipNamArr

For Each Rng In Range("a5:a16")
    If InStr(Rng.Value, nam) <> 0 Then
        Rows(Rng.Row).Select
    End If
Next Rng
Next nam


End Sub
Sub co()
Workbooks("PERSONVBA.xlsb").Sheets("回访打分表").Range("a1").Copy Workbooks("船队业务管理计划【201709】-小船总表.xls").Sheets("业务管理统计表").Range("c3")

End Sub
Sub se()
'
Sheets("航次增效报表").Range("j102").Copy
Sheets("航次增效报表").Range("j103:j113").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
   Application.CutCopyMode = False


End Sub
Sub 查找()
Attribute 查找.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 查找 宏
'

'
dh1 = "鼎衡1"
dh = "鼎衡１"
For Each Rng In Range("a4:a35")
Debug.Print Rng.Row
Debug.Print Rng.Value
Debug.Print Trim(Left(Rng.Value, 4)) = dh1
If Trim(Left(Rng.Value, 4)) = dh1 Then
Debug.Print "找到了："; Mid(Rng.Value, 1, Len(dh1)); "  "; Rng.Row
End If
Next Rng

End Sub
Sub 宏12()
Attribute 宏12.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏12 宏
'

'
    Windows("船队业务管理计划【201709】-小船总表.xls").Activate
    Sheets("时间管理统计表").Select
End Sub
Sub 宏15()
Attribute 宏15.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏15 宏
'

'
    Range("S5").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]-RC[-5]<0,0,RC[-1]-RC[-4])"
    Range("AC5").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]-RC[-5]<0,0,RC[-1]-RC[-4])"
    Range("AD5").Select
    Range("AM5").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]-RC[-5]<0,0,RC[-1]-RC[-4])"
    Range("AN5").Select
    ActiveWindow.SmallScroll ToRight:=9
    Range("AW5").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]-RC[-5]<0,0,RC[-1]-RC[-4])"
    Range("AX5").Select
    ActiveWindow.SmallScroll ToRight:=7
    Range("BG5").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-4]"
    Range("BH5").Select
End Sub
Sub fom()
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

End Sub
