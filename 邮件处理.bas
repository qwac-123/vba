Attribute VB_Name = "�ʼ�����"
Sub GetSanderAdressAndBody()  '//��÷����˵ĵ�ַ ������
    Dim Application As outlook.Application
    Dim myNamespace As Namespace
    Dim myFolder As MAPIFolder

    Dim Folder As MAPIFolder
    Dim iMail As outlook.MailItem


    Dim ExcelApp
    Set ExcelApp = GetObject("", "Excel.Application")
    Set wbk = ExcelApp.Workbooks.Open("f:/������.xlsx")

    Set wst = wbk.Sheets(1)

    Set Application = New outlook.Application
    Set myNamespace = Application.GetNamespace("MAPI")
    'Set myFolder = MyNameSpace.PickFolder
    Set myFolder = myNamespace.GetDefaultFolder(olFolderInbox)    '//����ռ����ļ���
    '// myNamespace.Folders.Count


    For i = 1 To myFolder.Folders.Count

        Set Folder = myFolder.Folders(i)

        For Each iMail In Folder.Items

            j = j + 1

            wst.Cells(j, 5) = iMail.ReceivedTime    '//�����ʼ�����ʱ��

            wst.Cells(j, 4) = Folder.Name    '//�����ļ�������

            wst.Cells(j, 1) = iMail.To    '//������

            wst.Cells(j, 2) = iMail.CC    '//������

            wst.Cells(j, 3) = iMail.Subject    '//����
        Next iMail

    Next

    wbk.Close True

    Set iMail = Nothing
    Set myFolder = Nothing
    Set myNamespace = Nothing
    Set Application = Nothing


End Sub
