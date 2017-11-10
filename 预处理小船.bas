Attribute VB_Name = "预处理小船"
Sub 清空小船总表()
Attribute 清空小船总表.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 清空小船总表 Macro
'

'
    Sheets("时间管理统计表").Select
    Range("J5").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=3
    Range("J5:K16").Select
    Selection.ClearContents
    
    
    Sheets("业务管理统计表").Select
    Range("B3:C14").Select
    Selection.ClearContents
    Sheets("航次增效报表").Select
    Range("B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("B4:E152").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("B4:N152").Select
    Selection.ClearContents
    Sheets("航次增效统计表").Select
    Range("F5:F16,D5:D16,B5:B16").Select
    Range("B16").Activate
    Selection.ClearContents
    Range("B5:G16").Select
    Range("G5").Activate
    Selection.ClearContents
    Sheets("业务管理计划核算表").Select
End Sub
Sub 复制公式格子()
Attribute 复制公式格子.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 复制公式格子 Macro
'

'
    Sheets("时间管理统计表").Select
    Range("A1:DH32").Select
    Range("D7").Activate
    Selection.ClearComments
    Range("J5").Select
    Range("J7").Select
    Application.FindFormat.Clear
    Application.FindFormat.NumberFormat = "0.00_);[红色](0.00)"
    With Application.FindFormat
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    With Application.FindFormat.Font
        .Name = "Arial Narrow"
        .FontStyle = "常规"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Application.FindFormat.Borders(xlLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Application.FindFormat.Borders(xlRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Application.FindFormat.Borders(xlTop)
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Application.FindFormat.Borders(xlBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Application.FindFormat.Borders(xlDiagonalDown).LineStyle = xlNone
    Application.FindFormat.Borders(xlDiagonalUp).LineStyle = xlNone
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ColorIndex = 2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Application.FindFormat.Locked = False
    Application.FindFormat.FormulaHidden = False
    Range("I5").Select
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 91
    ActiveWindow.ScrollColumn = 97
    ActiveWindow.ScrollColumn = 98
    ActiveWindow.ScrollColumn = 99
    ActiveWindow.ScrollColumn = 101
    ActiveWindow.ScrollColumn = 102
    ActiveWindow.ScrollColumn = 103
    ActiveWindow.ScrollColumn = 105
    ActiveWindow.ScrollColumn = 106
    ActiveWindow.SmallScroll Down:=6
    Range("I5:DG16").Select
    Selection.ClearContents
    Range("CV5").Select
    Application.FindFormat.Clear
    Application.FindFormat.NumberFormat = "0.00_);[红色](0.00)"
    With Application.FindFormat
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    With Application.FindFormat.Font
        .Name = "Arial Narrow"
        .FontStyle = "常规"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Application.FindFormat.Borders(xlLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Application.FindFormat.Borders(xlRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Application.FindFormat.Borders(xlTop)
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Application.FindFormat.Borders(xlBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Application.FindFormat.Borders(xlDiagonalDown).LineStyle = xlNone
    Application.FindFormat.Borders(xlDiagonalUp).LineStyle = xlNone
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ColorIndex = 19
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Application.FindFormat.Locked = True
    Application.FindFormat.FormulaHidden = False
    Range("J5").Select
    ActiveWindow.SmallScroll Down:=0
    Range("L5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("L16").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("L5").Select
    ActiveWindow.SmallScroll Down:=6
    Range("L5:DG16").Select

    Range("DA11").Select
End Sub
Sub 删除白格子()
Attribute 删除白格子.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 删除白格子 Macro
'

'
    Sheets("时间管理统计表").Select
    Range("J5").Select
    
    Application.FindFormat.Clear
    
    
    With Application.FindFormat.Font
        .Name = "宋体"
        .FontStyle = "Bold"
        .Size = 11
    End With
    Cells.Find(what:="", SearchFormat:=True).Activate
    
    
    
    
    With Application.FindFormat
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    With Application.FindFormat.Font
        .Name = "Arial Narrow"
        .FontStyle = "常规"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Application.FindFormat.Borders(xlLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Application.FindFormat.Borders(xlRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Application.FindFormat.Borders(xlTop)
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Application.FindFormat.Borders(xlBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Application.FindFormat.Borders(xlDiagonalDown).LineStyle = xlNone
    Application.FindFormat.Borders(xlDiagonalUp).LineStyle = xlNone
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ColorIndex = 2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Application.FindFormat.Locked = False
    Application.FindFormat.FormulaHidden = False
    Range("J5").Select
    Range("J5:DC16").Select
    ActiveWindow.SmallScroll Down:=-27
    Selection.ClearContents
End Sub
Sub 格式()
Attribute 格式.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 格式 Macro
'

'
    Range("J8").Select
    Application.FindFormat.Clear
    '5Application.FindFormat.NumberFormat = "0.00_);[红色](0.00)"
    With Application.FindFormat
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    With Application.FindFormat.Font
        .Name = "Arial Narrow"
        .FontStyle = "常规"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Application.FindFormat.Borders(xlLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Application.FindFormat.Borders(xlRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Application.FindFormat.Borders(xlTop)
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Application.FindFormat.Borders(xlBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Application.FindFormat.Borders(xlDiagonalDown).LineStyle = xlNone
    Application.FindFormat.Borders(xlDiagonalUp).LineStyle = xlNone
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ColorIndex = 2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Application.FindFormat.Locked = False
    Application.FindFormat.FormulaHidden = False
     'Range(Cells(5, 10), Cells(16, 20)).Find(what:="", SearchFormat:=True)
End Sub

Sub 空空()
Attribute 空空.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 空空 Macro
'

'
    Sheets("时间管理统计表").Select
    Range("J5:K5").Select
    ActiveWindow.SmallScroll Down:=9
    Range("J5:K16").Select
    ActiveWindow.SmallScroll Down:=-6
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=3
    Range("M16").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveWindow.SmallScroll Down:=-9
    Range("M5:M16").Select
    Range("M16").Activate
    Selection.ClearContents
    Range("O5").Select
    ActiveWindow.SmallScroll Down:=6
    Range("O5:R16").Select
    Selection.ClearContents
    Range("T16").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("T5:U16").Select
    Range("T16").Activate
    Selection.ClearContents
    Range("W5").Select
    ActiveWindow.SmallScroll Down:=6
    Range("W5:W16").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll ToRight:=8
    Range("AB16").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("Y5:AB16").Select
    Range("AB16").Activate
    Selection.ClearContents
    Range("AD5").Select
    ActiveWindow.SmallScroll Down:=9
    Range("AD5:AE16").Select
    Selection.ClearContents
    Range("AG5").Select
    ActiveWindow.SmallScroll Down:=9
    Range("AG5:AG16").Select
    Selection.ClearContents
    Range("AI16").Select
    ActiveWindow.SmallScroll ToRight:=5
    ActiveWindow.SmallScroll Down:=-12
    Range("AI5:AL16").Select
    Range("AI16").Activate
    Selection.ClearContents
    Range("AN5").Select
    ActiveWindow.SmallScroll ToRight:=4
    ActiveWindow.SmallScroll Down:=6
    Range("AN5:AO16").Select
    Selection.ClearContents
    Range("AQ16").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("AQ5:AQ16").Select
    Range("AQ16").Activate
    Selection.ClearContents
    Range("AS5").Select
    ActiveWindow.SmallScroll ToRight:=6
    ActiveWindow.SmallScroll Down:=9
    Range("AS5:AV16").Select
    Selection.ClearContents
    Range("AX16").Select
    ActiveWindow.SmallScroll ToRight:=5
    ActiveWindow.SmallScroll Down:=-12
    Range("AX5:AY16").Select
    Range("AX16").Activate
    Selection.ClearContents
    Range("BA5").Select
    ActiveWindow.SmallScroll Down:=9
    Range("BA5:BA16").Select
    Selection.ClearContents
    Range("BC16").Select
    ActiveWindow.SmallScroll Down:=-12
    ActiveWindow.SmallScroll ToRight:=6
    Range("BC5:BF16").Select
    Range("BC16").Activate
    ActiveWindow.SmallScroll Down:=9
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-12
    Range("BI5").Select
    ActiveWindow.SmallScroll ToRight:=8
    ActiveWindow.SmallScroll Down:=9
    Range("BI5:BZ16").Select
    Selection.ClearContents
    Range("CC16").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("CC5:CC16").Select
    Range("CC16").Activate
    Selection.ClearContents
    ActiveWindow.SmallScroll ToRight:=7
    Range("CE5").Select
    ActiveWindow.SmallScroll Down:=6
    Range("CE5:CE16").Select
    Selection.ClearContents
    Range("CG16").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("CG5:CG16").Select
    Range("CG16").Activate
    Selection.ClearContents
    Range("CI5").Select
    ActiveWindow.SmallScroll Down:=6
    Range("CI5:CI16").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll ToRight:=6
    Range("CK16").Select
    ActiveWindow.SmallScroll Down:=-15
    Range("CK5:CK16").Select
    Range("CK16").Activate
    Selection.ClearContents
    Range("CQ5").Select
    ActiveWindow.SmallScroll Down:=6
    Range("CQ5:CQ16").Select
    Selection.ClearContents
    Range("CS16").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("CS5:CS16").Select
    Range("CS16").Activate
    Selection.ClearContents
    Range("CU5").Select
    ActiveWindow.SmallScroll ToRight:=2
    ActiveWindow.LargeScroll ToRight:=0
    ActiveWindow.SmallScroll Down:=6
    Range("CU5:CU16").Select
    Selection.ClearContents
    Range("CW16").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("CW5:CW16").Select
    Range("CW16").Activate
    Selection.ClearContents
    ActiveWindow.SmallScroll ToRight:=5
    Range("CY5").Select
    ActiveWindow.SmallScroll Down:=6
    Range("CY5:CY16").Select
    Selection.ClearContents
    Range("DC16").Select
    ActiveWindow.SmallScroll Down:=-15
    Range("DC5:DC16").Select
    Range("DC16").Activate
    Selection.ClearContents
    ActiveWindow.SmallScroll ToRight:=6
    Range("DH5").Select
    ActiveWindow.SmallScroll Down:=9
    Range("DH5:DH16").Select
    Selection.ClearContents
    Range("A1:DH32").Select
    Range("DH5").Activate
    Selection.ClearComments
    Sheets("业务管理统计表").Select
    Range("C14").Select
    ActiveWindow.SmallScroll Down:=-6
    Range("B3:C14").Select
    Range("C14").Activate
    Selection.ClearContents
    Sheets("航次增效报表").Select
    Range("A4:A13").Select
    ActiveWindow.SmallScroll Down:=-6
    Range("C4").Select
    ActiveWindow.SmallScroll Down:=135
    Range("C4:N153").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-15
    Selection.ClearComments
    Sheets("航次增效统计表").Select
    Range("B5:G16").Select
    Selection.ClearContents
    Sheets("时间管理统计表").Select
    Range("K4").Select
End Sub
