Attribute VB_Name = "模块6"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Columns("B:B").Select
    Selection.FormulaR1C1 = "=+" & Chr(10) & "Bunkered this voyage"
    Application.FormulaBarHeight = 5
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Columns("B:B").Select
    Selection.FormulaR1C1 = "+"
    With Selection.Characters(start:=1, Length:=1).Font
        .Name = "宋体"
        .FontStyle = "常规"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = 2
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    Columns("B:B").Select
    Range("B13").Activate
    Selection.FormulaR1C1 = "end"
    With Selection.Characters(start:=1, Length:=3).Font
        .Name = "宋体"
        .FontStyle = "常规"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = 2
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Cells.Select
    Range("A13").Activate
    Selection.Rows.AutoFit
    Selection.RowHeight = 6
    Selection.RowHeight = 20
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    Columns("A:A").Select
    Selection.Replace What:="鼎衡", Replacement:="DH", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="轮", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="：", Replacement:=":", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    Range("D17").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("E17").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("D17").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("E17").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub cl()
ActiveCell.Interior.Pattern = xlNone
ActiveCell.Interior.color = 65535
End Sub
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    Range("F4").Select
    ActiveWindow.SmallScroll Down:=3
    Range("F4:F16").Select
    Selection.FormulaR1C1 = _
        "=IF(RC[1]<>"""",""开往""&MID(RC[1],5,3),IF(RC[2]<>"""",""锚泊""&MID(RC[2],5,3),IF(COUNT(FIND(""靠泊"",RC[5])),IF(SUM(ISNUMBER(FIND({""张家港"",""连云港"",""鲅鱼圈"",""仙人岛""},RC[5]))*1),MID(RC[5],FIND(""靠泊"",RC[5]),5),MID(RC[5],FIND(""靠泊"",RC[5]),4)),RC[6]&""完货"")))"
End Sub
