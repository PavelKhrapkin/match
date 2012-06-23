VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SFaccountForm 
   Caption         =   "Выбор SF предприятия"
   ClientHeight    =   10530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   22755
   OleObjectBlob   =   "SFaccountForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SFaccountForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub accntChoice_Click()
    SFaccountForm.TextBox1 = Str(Me.accntChoice.ListIndex + 1)    ' items считаем нумерованными с 1
End Sub

Private Sub createbutton_Click()
    Me.TextBox1 = "create"
    SFaccountForm.Hide

End Sub


Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub OKButton_Click()
    SFaccountForm.Hide
End Sub
Private Sub contButton_Click()
    Me.TextBox1 = "cont"
    SFaccountForm.Hide

End Sub
Private Sub ExitButton_Click()
    Me.TextBox1 = "exit"
    SFaccountForm.Hide
End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub accntChoice_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SFaccountForm.TextBox1 = Str(Me.accntChoice.ListIndex + 1)    ' items считаем нумерованными с 1
    SFaccountForm.Hide
End Sub

Private Sub UserForm_Click()

End Sub
