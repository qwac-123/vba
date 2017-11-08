Attribute VB_Name = "小船公式模块"
Sub kongkong()
Call 空空
End Sub
Sub 小船总表一键搞定()
Attribute 小船总表一键搞定.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 小船一键
'

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False

x = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", _
       Title:="选择小船表", MultiSelect:=True) '选择要被合并的簿
If TypeName(x) <> "Variant()" Then
    Exit Sub
End If
'shipNamArr = Array("鼎衡1", "鼎衡10", "鼎衡15", "鼎衡16", "鼎衡17", "鼎衡18", "鼎衡2", "鼎衡3", "建兴32", "鼎衡5", "鼎衡7", "鼎衡9")
lujin = Mid(x(1), 1, 24)

For Each x1 In x
    Debug.Print Mid(x1, 14, InStr(14, x1, ".") - 14)
    shipNam = Mid(x1, 14, InStr(14, x1, ".") - 14)
    zong = ActiveWorkbook.Name
    shipBk = Workbooks.Open(x1).Name
    Workbooks(zong).Activate
    For Each biaobiao In Sheets
        biaobiao.Cells.Replace "１", "1"
    Next
    Workbooks(shipBk).Activate
    For Each biaobiao In Sheets
        biaobiao.Cells.Replace "１", "1"
        Debug.Print biaobiao.Name
    Next
    Workbooks(zong).Activate
    For Each Rng In Workbooks(zong).Sheets("时间管理统计表").Range("a5:a22")
        If Len(shipNam) = 3 Then
            If Mid(Rng.Value, 4, 1) = Chr(10) Then
                If Trim(Left(Rng.Value, 3)) = shipNam Then
                    zongshiRow = Rng.Row
                    Exit For
                End If
            End If
        Else
            If Trim(Left(Rng.Value, 4)) = shipNam Then
                zongshiRow = Rng.Row
                Exit For
            End If
        End If
    Next Rng
    For Each Rng In Workbooks(zong).Sheets("业务管理统计表").Range("a2:a19")
        If Rng.Value = shipNam Then
            zongguanrow = Rng.Row
            Exit For
        End If
    Next Rng
    For Each Rng In Workbooks(zong).Sheets("航次增效统计表").Range("a4:a19")
        If Rng.Value = shipNam Then
            zongtongrow = Rng.Row
            Exit For
        End If
    Next Rng

 '总表的航次增效报表要根据分表的数量来调整，在后面

    For Each Rng In Workbooks(shipBk).Sheets("时间管理统计表").Range("a2:a19")
        If Len(shipNam) = 3 Then
            If Mid(Rng.Value, 4, 1) = Chr(10) Then
                If Trim(Left(Rng.Value, 3)) = shipNam Then
                    Debug.Print shipNam, "时间管理统计表行号", Rng.Row
                    Workbooks(shipBk).Sheets("时间管理统计表").Rows(Rng.Row).Copy Workbooks(zong).Sheets("时间管理统计表").Rows(zongshiRow)
                    Exit For
                End If
            End If
        Else
            If Trim(Left(Rng.Value, 4)) = shipNam Then
                Workbooks(shipBk).Sheets("时间管理统计表").Rows(Rng.Row).Copy Workbooks(zong).Sheets("时间管理统计表").Rows(zongshiRow)
                Exit For
            End If
        End If
    Next Rng



    For Each Rng In Workbooks(shipBk).Sheets("业务管理统计表").Range("a2:f2") '找到多航次营运那一列
        If InStr(Rng.Value, "多航次营运") Then
            colduo = Rng.Column
            Exit For
        End If
    Next Rng
    For Each Rng In Workbooks(shipBk).Sheets("业务管理统计表").Range("a2:a18") '找到本船所在的行
        If Rng.Value = shipNam Then
            If Workbooks(shipBk).Sheets("业务管理统计表").Cells(Rng.Row, colduo) = "" Then
                Workbooks(zong).Sheets("业务管理统计表").Cells(zongguanrow, 2) = "原表空"
            Else
                Workbooks(shipBk).Sheets("业务管理统计表").Cells(Rng.Row, colduo).Copy Workbooks(zong).Sheets("业务管理统计表").Cells(zongguanrow, 2)
                Exit For
            End If
        End If
    Next Rng
guanrow:
    For Each Rng In Workbooks(shipBk).Sheets("航次增效报表").Range("a2:a180") '分表航次增效报表找本船
        i = Rng.Row
        
            If Rng.Text = shipNam Then
                rySize = Rng.MergeArea.Rows.Count
                rend = i + rySize - 1
                For e = i To rend
                    If Workbooks(shipBk).Sheets("航次增效报表").Cells(e, 5) = "" Then
                        rMaxe = e - 1
                        Exit For
                    End If
                Next e
ii:
                For ii = i To rend
                    If Workbooks(shipBk).Sheets("航次增效报表").Cells(ii, 9) = "" Then
                        rMaxii = ii - 1
                        Exit For
                    End If
                Next ii
m:
                For m = i To rend
                    If Workbooks(shipBk).Sheets("航次增效报表").Cells(m, 13) = "" Then
                        rMaxm = m - 1
                        Exit For
                    End If
                Next m
rma:
                rMax = WorksheetFunction.Max(rMaxe, rMaxii, rMaxm)
                MsgBox rMax
                rMin = i
                Exit For
            End If
    Next Rng
over:
    For Each Rng In Workbooks(zong).Sheets("航次增效报表").Range("a2:a180") '总表增效报表找位置
        i = Rng.Row
        If Rng.Text = shipNam Then
            rzSize = Rng.MergeArea.Rows.Count
            charu = rySize - rzSize
                If charu > 0 Then
                    For konghang = 1 To charu
                        Workbooks(zong).Sheets("航次增效报表").Cells(i + 1, 2).EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
                    Next konghang
                End If
            Workbooks(shipBk).Sheets("航次增效报表").Activate
            Workbooks(shipBk).Sheets("航次增效报表").Range(Cells(rMin, 2), Cells(rMax, 13)).Copy
            Workbooks(zong).Sheets("航次增效报表").Activate
            Cells(i, 2).Activate
            ActiveSheet.Paste
            Cells(i, 2).Activate
            
        Exit For
        End If
    Next Rng
hangcitongji:
Workbooks(shipBk).Sheets("航次增效统计表").Activate
    For Each Rng In Workbooks(shipBk).Sheets("航次增效统计表").Range("a3:a18")
        If Rng.Value = shipNam Then

            If Not Workbooks(shipBk).Sheets("航次增效统计表").Cells(Rng.Row, 2) > 0 Then
                Workbooks(zong).Sheets("航次增效统计表").Cells(zongtongrow, 2) = "空"
            Else
                Workbooks(shipBk).Sheets("航次增效统计表").Cells(Rng.Row, 2).Copy Workbooks(zong).Sheets("航次增效统计表").Cells(zongtongrow, 2)
            End If
            
            If Not Workbooks(shipBk).Sheets("航次增效统计表").Cells(Rng.Row, 4) > 0 Then
                Workbooks(zong).Sheets("航次增效统计表").Cells(zongtongrow, 4) = "空"
            Else
                Workbooks(shipBk).Sheets("航次增效统计表").Cells(Rng.Row, 4).Copy Workbooks(zong).Sheets("航次增效统计表").Cells(zongtongrow, 4)
            End If
            
            If Not Workbooks(shipBk).Sheets("航次增效统计表").Cells(Rng.Row, 6) > 0 Then
                Workbooks(zong).Sheets("航次增效统计表").Cells(zongtongrow, 6) = "空"
            Else
                Workbooks(shipBk).Sheets("航次增效统计表").Cells(Rng.Row, 6).Copy Workbooks(zong).Sheets("航次增效统计表").Cells(zongtongrow, 6)
            End If
            
            Exit For
        End If
    Next Rng
tongrow:
Windows(shipBk).Close
Next x1
'整理函数
Windows(zong).Activate
Sheets("时间管理统计表").Select
zongshiRow = 5
Range("l" & CStr(zongshiRow)).FormulaR1C1 = "=IF(RC[-1]=0,0,RC[-2]/RC[-1])"
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
Range("BG" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-1]-RC[-4]"
Range("BH" & CStr(zongshiRow)).FormulaR1C1 = "=RC[-41]+RC[-31]+RC[-21]+RC[-11]+RC[-1]"
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

zongweiRow = InputBox("最后一行", "请输入最后一条船的行号")

Cells(zongshiRow, 12).AutoFill Destination:=Range(Cells(zongshiRow, 12), Cells(zongweiRow, 12)), Type:=xlFillDefault
Cells(zongshiRow, 14).AutoFill Destination:=Range(Cells(zongshiRow, 14), Cells(zongweiRow, 14)), Type:=xlFillDefault
Cells(zongshiRow, 19).AutoFill Destination:=Range(Cells(zongshiRow, 19), Cells(zongweiRow, 19)), Type:=xlFillDefault
Cells(zongshiRow, 22).AutoFill Destination:=Range(Cells(zongshiRow, 22), Cells(zongweiRow, 22)), Type:=xlFillDefault
Cells(zongshiRow, 24).AutoFill Destination:=Range(Cells(zongshiRow, 24), Cells(zongweiRow, 24)), Type:=xlFillDefault
Cells(zongshiRow, 29).AutoFill Destination:=Range(Cells(zongshiRow, 29), Cells(zongweiRow, 29)), Type:=xlFillDefault
Cells(zongshiRow, 32).AutoFill Destination:=Range(Cells(zongshiRow, 32), Cells(zongweiRow, 32)), Type:=xlFillDefault
Cells(zongshiRow, 34).AutoFill Destination:=Range(Cells(zongshiRow, 34), Cells(zongweiRow, 34)), Type:=xlFillDefault
Cells(zongshiRow, 39).AutoFill Destination:=Range(Cells(zongshiRow, 39), Cells(zongweiRow, 39)), Type:=xlFillDefault
Cells(zongshiRow, 42).AutoFill Destination:=Range(Cells(zongshiRow, 42), Cells(zongweiRow, 42)), Type:=xlFillDefault
Cells(zongshiRow, 44).AutoFill Destination:=Range(Cells(zongshiRow, 44), Cells(zongweiRow, 44)), Type:=xlFillDefault
Cells(zongshiRow, 49).AutoFill Destination:=Range(Cells(zongshiRow, 49), Cells(zongweiRow, 49)), Type:=xlFillDefault
Cells(zongshiRow, 52).AutoFill Destination:=Range(Cells(zongshiRow, 52), Cells(zongweiRow, 52)), Type:=xlFillDefault
Cells(zongshiRow, 54).AutoFill Destination:=Range(Cells(zongshiRow, 54), Cells(zongweiRow, 54)), Type:=xlFillDefault
Cells(zongshiRow, 59).AutoFill Destination:=Range(Cells(zongshiRow, 59), Cells(zongweiRow, 59)), Type:=xlFillDefault
Cells(zongshiRow, 60).AutoFill Destination:=Range(Cells(zongshiRow, 60), Cells(zongweiRow, 60)), Type:=xlFillDefault
Cells(zongshiRow, 82).AutoFill Destination:=Range(Cells(zongshiRow, 82), Cells(zongweiRow, 82)), Type:=xlFillDefault
Cells(zongshiRow, 84).AutoFill Destination:=Range(Cells(zongshiRow, 84), Cells(zongweiRow, 84)), Type:=xlFillDefault
Cells(zongshiRow, 86).AutoFill Destination:=Range(Cells(zongshiRow, 86), Cells(zongweiRow, 86)), Type:=xlFillDefault
Cells(zongshiRow, 88).AutoFill Destination:=Range(Cells(zongshiRow, 88), Cells(zongweiRow, 88)), Type:=xlFillDefault
Cells(zongshiRow, 90).AutoFill Destination:=Range(Cells(zongshiRow, 90), Cells(zongweiRow, 90)), Type:=xlFillDefault
Cells(zongshiRow, 96).AutoFill Destination:=Range(Cells(zongshiRow, 96), Cells(zongweiRow, 96)), Type:=xlFillDefault
Cells(zongshiRow, 98).AutoFill Destination:=Range(Cells(zongshiRow, 98), Cells(zongweiRow, 98)), Type:=xlFillDefault
Cells(zongshiRow, 100).AutoFill Destination:=Range(Cells(zongshiRow, 100), Cells(zongweiRow, 100)), Type:=xlFillDefault
Cells(zongshiRow, 102).AutoFill Destination:=Range(Cells(zongshiRow, 102), Cells(zongweiRow, 102)), Type:=xlFillDefault
Cells(zongshiRow, 104).AutoFill Destination:=Range(Cells(zongshiRow, 104), Cells(zongweiRow, 104)), Type:=xlFillDefault
Cells(zongshiRow, 105).AutoFill Destination:=Range(Cells(zongshiRow, 105), Cells(zongweiRow, 105)), Type:=xlFillDefault
Cells(zongshiRow, 109).AutoFill Destination:=Range(Cells(zongshiRow, 109), Cells(zongweiRow, 109)), Type:=xlFillDefault
Cells(zongshiRow, 111).AutoFill Destination:=Range(Cells(zongshiRow, 111), Cells(zongweiRow, 111)), Type:=xlFillDefault
'最后统一格式
    Rows("16").Copy
    Rows("5:16").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Cells(5, 1).Select

    Sheets("业务管理统计表").Select
    Rows("14").Copy
    Rows("3:14").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Cells(3, 1).Select

    Sheets("航次增效统计表").Select
    Rows("16").Copy
    Rows("5:16").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Cells(5, 1).Select

' 改时间 Macro
Sheets("时间管理统计表").Range("B1:P1") = Format(Date, "船舶月度时间管理统计及奖金计算表（yyyy年mm月）　")
Sheets("航次增效报表").Range("C1:N1") = Format(Date, "船舶月度节能增效报表(yyyy年mm月）")
Sheets("航次增效统计表").Range("B1:I1") = Format(Date, "船舶月度节能增效及奖金计算表(yyyy年mm月）")
Sheets("业务管理计划核算表").Range("M1") = Format(Date, "yyyy年mm月")

Sheets("时间管理统计表").Select

Application.ScreenUpdating = 1
Application.DisplayAlerts = 1
End Sub
Sub SBUS()
zongshiRow = 5
zongweiRow = InputBox("最后一行", "请输入最后一条船的行号", 16)
Debug.Print zongweiRow
Cells(zongshiRow, 12).AutoFill Destination:=Range(Cells(zongshiRow, 12), Cells(zongweiRow, 12)), Type:=xlFillDefault
Cells(zongshiRow, 14).AutoFill Destination:=Range(Cells(zongshiRow, 14), Cells(zongweiRow, 14)), Type:=xlFillDefault
Cells(zongshiRow, 19).AutoFill Destination:=Range(Cells(zongshiRow, 19), Cells(zongweiRow, 19)), Type:=xlFillDefault
Cells(zongshiRow, 22).AutoFill Destination:=Range(Cells(zongshiRow, 22), Cells(zongweiRow, 22)), Type:=xlFillDefault
Cells(zongshiRow, 24).AutoFill Destination:=Range(Cells(zongshiRow, 24), Cells(zongweiRow, 24)), Type:=xlFillDefault
Cells(zongshiRow, 29).AutoFill Destination:=Range(Cells(zongshiRow, 29), Cells(zongweiRow, 29)), Type:=xlFillDefault
Cells(zongshiRow, 32).AutoFill Destination:=Range(Cells(zongshiRow, 32), Cells(zongweiRow, 32)), Type:=xlFillDefault
Cells(zongshiRow, 34).AutoFill Destination:=Range(Cells(zongshiRow, 34), Cells(zongweiRow, 34)), Type:=xlFillDefault
Cells(zongshiRow, 39).AutoFill Destination:=Range(Cells(zongshiRow, 39), Cells(zongweiRow, 39)), Type:=xlFillDefault
Cells(zongshiRow, 42).AutoFill Destination:=Range(Cells(zongshiRow, 42), Cells(zongweiRow, 42)), Type:=xlFillDefault
Cells(zongshiRow, 44).AutoFill Destination:=Range(Cells(zongshiRow, 44), Cells(zongweiRow, 44)), Type:=xlFillDefault
Cells(zongshiRow, 49).AutoFill Destination:=Range(Cells(zongshiRow, 49), Cells(zongweiRow, 49)), Type:=xlFillDefault
Cells(zongshiRow, 52).AutoFill Destination:=Range(Cells(zongshiRow, 52), Cells(zongweiRow, 52)), Type:=xlFillDefault
Cells(zongshiRow, 54).AutoFill Destination:=Range(Cells(zongshiRow, 54), Cells(zongweiRow, 54)), Type:=xlFillDefault
Cells(zongshiRow, 59).AutoFill Destination:=Range(Cells(zongshiRow, 59), Cells(zongweiRow, 59)), Type:=xlFillDefault
Cells(zongshiRow, 60).AutoFill Destination:=Range(Cells(zongshiRow, 60), Cells(zongweiRow, 60)), Type:=xlFillDefault
Cells(zongshiRow, 82).AutoFill Destination:=Range(Cells(zongshiRow, 82), Cells(zongweiRow, 82)), Type:=xlFillDefault
Cells(zongshiRow, 84).AutoFill Destination:=Range(Cells(zongshiRow, 84), Cells(zongweiRow, 84)), Type:=xlFillDefault
Cells(zongshiRow, 86).AutoFill Destination:=Range(Cells(zongshiRow, 86), Cells(zongweiRow, 86)), Type:=xlFillDefault
Cells(zongshiRow, 88).AutoFill Destination:=Range(Cells(zongshiRow, 88), Cells(zongweiRow, 88)), Type:=xlFillDefault
Cells(zongshiRow, 90).AutoFill Destination:=Range(Cells(zongshiRow, 90), Cells(zongweiRow, 90)), Type:=xlFillDefault
Cells(zongshiRow, 96).AutoFill Destination:=Range(Cells(zongshiRow, 96), Cells(zongweiRow, 96)), Type:=xlFillDefault
Cells(zongshiRow, 98).AutoFill Destination:=Range(Cells(zongshiRow, 98), Cells(zongweiRow, 98)), Type:=xlFillDefault
Cells(zongshiRow, 100).AutoFill Destination:=Range(Cells(zongshiRow, 100), Cells(zongweiRow, 100)), Type:=xlFillDefault
Cells(zongshiRow, 102).AutoFill Destination:=Range(Cells(zongshiRow, 102), Cells(zongweiRow, 102)), Type:=xlFillDefault
Cells(zongshiRow, 104).AutoFill Destination:=Range(Cells(zongshiRow, 104), Cells(zongweiRow, 104)), Type:=xlFillDefault
Cells(zongshiRow, 105).AutoFill Destination:=Range(Cells(zongshiRow, 105), Cells(zongweiRow, 105)), Type:=xlFillDefault
Cells(zongshiRow, 109).AutoFill Destination:=Range(Cells(zongshiRow, 109), Cells(zongweiRow, 109)), Type:=xlFillDefault
Cells(zongshiRow, 111).AutoFill Destination:=Range(Cells(zongshiRow, 111), Cells(zongweiRow, 111)), Type:=xlFillDefault

End Sub

Sub orr()
For Each rn In Range("f4:f15")
If rn Then
Debug.Print "格子里是"; rn; "row"; rn.Row
Else
Debug.Print "空"; "row"; rn.Row
End If

Next
End Sub
Sub 空空()
'
' 空空 Macro
'

'
Application.ScreenUpdating = 0
Application.DisplayAlerts = 0
    Sheets("时间管理统计表").Select
    Range("J5:K5").Select
    Range("J5:K16").Select
    Selection.ClearContents
    Range("M16").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M16").Select
    Range("M16").Activate
    Selection.ClearContents
    Range("O5").Select
    Range("O5:R16").Select
    Selection.ClearContents
    Range("T16").Select
    Range("T5:U16").Select
    Range("T16").Activate
    Selection.ClearContents
    Range("W5").Select
    Range("W5:W16").Select
    Selection.ClearContents
    Range("AB16").Select
    Range("Y5:AB16").Select
    Range("AB16").Activate
    Selection.ClearContents
    Range("AD5").Select
    Range("AD5:AE16").Select
    Selection.ClearContents
    Range("AG5").Select
    Range("AG5:AG16").Select
    Selection.ClearContents
    Range("AI16").Select
    Range("AI5:AL16").Select
    Range("AI16").Activate
    Selection.ClearContents
    Range("AN5").Select
    Range("AN5:AO16").Select
    Selection.ClearContents
    Range("AQ16").Select
    Range("AQ5:AQ16").Select
    Range("AQ16").Activate
    Selection.ClearContents
    Range("AS5").Select
    Range("AS5:AV16").Select
    Selection.ClearContents
    Range("AX16").Select
    Range("AX5:AY16").Select
    Range("AX16").Activate
    Selection.ClearContents
    Range("BA5").Select
    Range("BA5:BA16").Select
    Selection.ClearContents
    Range("BC16").Select
    Range("BC5:BF16").Select
    Range("BC16").Activate
    Selection.ClearContents
    Range("BI5").Select
    Range("BI5:BZ16").Select
    Selection.ClearContents
    Range("CC16").Select
    Range("CC5:CC16").Select
    Range("CC16").Activate
    Selection.ClearContents
    Range("CE5").Select
    Range("CE5:CE16").Select
    Selection.ClearContents
    Range("CG16").Select
    Range("CG5:CG16").Select
    Range("CG16").Activate
    Selection.ClearContents
    Range("CI5").Select
    Range("CI5:CI16").Select
    Selection.ClearContents
    Range("CK16").Select
    Range("CK5:CK16").Select
    Range("CK16").Activate
    Selection.ClearContents
    Range("CQ5").Select
    Range("CQ5:CQ16").Select
    Selection.ClearContents
    Range("CS16").Select
    Range("CS5:CS16").Select
    Range("CS16").Activate
    Selection.ClearContents
    Range("CU5").Select
    Range("CU5:CU16").Select
    Selection.ClearContents
    Range("CW16").Select
    Range("CW5:CW16").Select
    Range("CW16").Activate
    Selection.ClearContents
    Range("CY5").Select
    Range("CY5:CY16").Select
    Selection.ClearContents
    Range("DC16").Select
    Range("DC5:DC16").Select
    Range("DC16").Activate
    Selection.ClearContents
    Range("DH5").Select
    Range("DH5:DH16").Select
    Selection.ClearContents
    Range("A1:DH32").Select
    Range("DH5").Activate
    Selection.ClearComments
    Sheets("业务管理统计表").Select
    Range("C14").Select
    Range("B3:C14").Select
    Range("C14").Activate
    Selection.ClearContents
    Sheets("航次增效报表").Select
    Range("A4:A13").Select
    Range("C4").Select
    Range("C4:N153").Select
    Selection.ClearContents
    Selection.ClearComments
    Sheets("航次增效统计表").Select
    Range("B5:G16").Select
    Selection.ClearContents
    Sheets("时间管理统计表").Select
    Range("K4").Select
Application.ScreenUpdating = 1
Application.DisplayAlerts = 1

End Sub
