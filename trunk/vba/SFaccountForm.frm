VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SFaccountForm 
   Caption         =   "Выбор SF предприятия для MS лида"
   ClientHeight    =   10620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12285
   OleObjectBlob   =   "SFaccountForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SFaccountForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ListBox1_Click()

End Sub

Private Sub createbutton_Click()
    Me.TextBox1 = "create"
    SFaccountForm.Hide

End Sub


Private Sub OKButton_Click()
    SFaccountForm.Hide

End Sub
Private Sub contButton_Click()
    Me.TextBox1 = "cont"
    SFaccountForm.Hide

End Sub
Private Sub exitButton_Click()
    Me.TextBox1 = "exit"
    SFaccountForm.Hide

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Click()

End Sub

