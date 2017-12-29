'v2 这次是把所有航次报表的内容集合到一起
Function 增加表并改名()
shipLonNamArr = Array("鼎衡1", "鼎衡10", "鼎衡15", "鼎衡16", "鼎衡17*", "鼎衡18*", "鼎衡2", "鼎衡3", "建兴32", "鼎衡5", "鼎衡7", "鼎衡9", "鼎衡A", "鼎衡E", "天使1", "天使2", "天使3", "天使11")
shipNamArr = Array("DH1", "DH10", "DH15", "DH16", "DH17", "DH18", "DH2", "DH3", "JX32", "DH5", "DH7", "DH9", "DHA", "DHE", "AG1", "AG2", "AG3", "AG11")
For i = 0 To UBound(shipNamArr)
    On Error Resume Next
    Sheets(i + 1).Name = shipNamArr(i)
    Debug.Print Err.Number
    If Err.Number = 9 Then
        Sheets.Add after:=Sheets(i)
        Sheets(i + 1).Name = shipNamArr(i)
    End If
Next i
End Function
Function 整合整合()
shipLongNameArr = Array("鼎衡1", "鼎衡2", "鼎衡3", "鼎衡5", "鼎衡9", "鼎衡10", "鼎衡15", "鼎衡16", "鼎衡17*", "鼎衡18*", "鼎衡7", "建兴32", "鼎衡A", "鼎衡E", "天使1", "天使2", "天使3", "天使11")
shipNameArr = Array("DH1", "DH2", "DH3", "DH5", "DH9", "DH10", "DH15", "DH16", "DH17", "DH18", "DH7", "JX32", "DHA", "DHE", "AG1", "AG2", "AG3", "AG11")

For isht = 2 To UBound(shipNameArr) + 1
    Set sht = Sheets(isht)
    shipNam = shipLongNameArr(isht - 1)
    shipNamShort = sht.Name
    
    If sht.[b3].Value <> "" Then
        startRow = sht.[b3].End(xlDown).Row + 1
    ElseIf sht.[b2].Value = "" Then
        startRow = 3
    Else
        startRow = 2
    End If
    endRow = sht.[a1].End(xlDown).Row
    voy = sht.Cells(startRow, 1).Value
    Call 获取该船所有大于等于所需航次的文件(shipNam, voy)
    For ro = startRow To endRow
        voy = sht.Cells(ro, 1).Value
        oilDir = "\\192.168.0.223\航运在线\10、油料管理部\航次报表\" & shipNam & "\2017年\" & "*燃*航次报表*" & voy & "*.xls?"
        voyDir = "\\192.168.0.223\航运在线\10、油料管理部\航次报表\" & shipNam & "\2017年\" & "*航次报表*" & voy & "*.xls?"
        
        If Dir(oilDir) = Dir(voyDir) Then
            
        voy = "V" & voy
        If Len(Dir(voyDir)) > 0 Then '如果文件存在
            
            oilNam = Dir(oilDir)
            voyNam = Dir(voyDir)
            filePath = "\\192.168.0.223\航运在线\10、油料管理部\航次报表\" & shipNam & "\2017年\"
            Set oilW = Workbooks.Open(fileName:=oilDir)
            Set voyW = Workbooks.Open(fileName:=oilDir)
            filePath = voyW.Path & "\"
            
            oilW.SaveAs filePath & shipNamShort & "燃润料航次报表" & voy & ".xlsx"
            voyW.SaveAs filePath & shipNamShort & "航次报表" & voy & ".xlsx"
            
            oilW.Close False
            voyW.Close False
            Kill 1
            Kill 2
        End If
    Next ro
Next isht
End Function
Function 获取该船所有大于等于所需航次的文件(shipNam, voy)
i = 0
fileDir = "\\192.168.0.223\航运在线\10、油料管理部\航次报表\" & shipNam & "\2017年\"
wjm = Dir(fileDir)
Do
If 提取航次号(wjm) >= voy Then
    arFiles(i) = wjm
    i = i + 1
End If
Loop While (1)
End Function
Function w()
Call 获取该船所有大于等于所需航次的文件("鼎衡2", "1743")
Call 选取该航次文件("鼎衡2", "1743")
End Function
Function 选取该航次文件(shipNam, voy)
fileDir = "\\192.168.0.223\航运在线\10、油料管理部\航次报表\" & shipNam & "\2017年\"
wjm = Dir(fileDir)
Do
If wjm Like "*燃*" & voy Then
    oilNam = wjm
    wjm = Dir
ElseIf wjm Like "*航次*" & voy Then
    voyNam = Dir
Else
    wjm = Dir
End If
Loop While (oilNam = "" Or voyNam = "" Or oilNam = voyNam)
Debug.Print oilNam <> "" And voyNam <> "" And oilNam <> voyNam
End Function
Function 提取航次号(fileName)
fileName = Application.Asc(fileName)
Set regVoy = CreateObject("vbscript.regexp")
regVoy.Pattern = "\d\d\d\d"
Set voySet = regVoy.Execute(fileName) '  Execute方法返回的集合对象mh,有两个属性:
For Each voy In voySet
    提取数字 = voy.Value
Next
End Function
Function 提取数字(rngValue)
rngValue = Application.Asc(rngValue)
Set regNum = CreateObject("vbscript.regexp")
regNum.Pattern = "\d+.\d+"
Set mh = regNum.Execute(rngValue) '  Execute方法返回的集合对象mh,有两个属性:
For Each mhk In mh
    提取数字 = mhk.Value
Next
End Function
