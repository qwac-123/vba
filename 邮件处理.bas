Attribute VB_Name = "邮件处理"
Sub GetSanderAdressAndBody()  '//获得发件人的地址 和正文
    Dim Application As outlook.Application
    Dim myNamespace As Namespace
    Dim myFolder As MAPIFolder

    Dim Folder As MAPIFolder
    Dim iMail As outlook.MailItem


    Dim ExcelApp
    Set ExcelApp = GetObject("", "Excel.Application")
    Set wbk = ExcelApp.Workbooks.Open("f:/测试中.xlsx")

    Set wst = wbk.Sheets(1)

    Set Application = New outlook.Application
    Set myNamespace = Application.GetNamespace("MAPI")
    'Set myFolder = MyNameSpace.PickFolder
    Set myFolder = myNamespace.GetDefaultFolder(olFolderInbox)    '//获得收件箱文件夹
    '// myNamespace.Folders.Count


    For i = 1 To myFolder.Folders.Count

        Set Folder = myFolder.Folders(i)

        For Each iMail In Folder.Items

            j = j + 1

            wst.Cells(j, 5) = iMail.ReceivedTime    '//接收邮件日期时间

            wst.Cells(j, 4) = Folder.Name    '//所在文件夹名称

            wst.Cells(j, 1) = iMail.To    '//发件人

            wst.Cells(j, 2) = iMail.CC    '//抄送人

            wst.Cells(j, 3) = iMail.Subject    '//正文
        Next iMail

    Next

    wbk.Close True

    Set iMail = Nothing
    Set myFolder = Nothing
    Set myNamespace = Nothing
    Set Application = Nothing


End Sub
