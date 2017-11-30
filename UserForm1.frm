VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4512
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7164
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub ScrollBar1_Change()
maxVl = 1000
minvl = 0
ScrollBar1.Min = maxVl
ScrollBar1.Max = minvl
ScrollBar1.SmallChange = 10
ScrollBar1.LargeChange = 100

If ScrollBar1.Value = maxVl Then
    Label1.Caption = "All in"
Else
    Label1.Caption = ScrollBar1.Value
End If
End Sub

Private Sub TextBox1_Change()
TextBox1.Text = 1
End Sub
