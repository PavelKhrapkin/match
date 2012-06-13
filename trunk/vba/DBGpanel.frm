VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DBGpanel 
   Caption         =   "Debug (Ctrl/Shift/Q)"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "DBGpanel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DBGpanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-------------------------------------------------------
' Debug3pasForm - отладочная панель для модуля 3PASS
'   1.3.2012

Option Explicit
Private Sub CheckContrButton_Click()
    DBGpanel.Hide
    ContrPass
End Sub
Private Sub CheckingButton_Click()
    DBGpanel.Hide
'    CheckingForm.Show
    CheckFofmOutput     ' вывод CheckingForm с формированием списка Продавцов
End Sub
Private Sub Debug3passForm_Click()
    DBGpanel.Hide
    Debug3pass.Show
    End
End Sub
Private Sub NewDogButton_Click()
    DBGpanel.Hide
    NewContractDL       ' создание новых Договоров по листу Договоров (3)
End Sub

Private Sub P_HandlingButton_Click()
    DBGpanel.Hide
    PaidHandling        ' проход по листу Платежей (1)
End Sub
