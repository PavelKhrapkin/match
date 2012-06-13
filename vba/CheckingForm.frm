VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CheckingForm 
   Caption         =   "Checking"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2025
   OleObjectBlob   =   "CheckingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CheckingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------
' CheckingForm - проверка Match
'   17.2.2012

Option Explicit
Private Sub AllSalesButton_Click()
    CheckingForm.Hide
    CheckPaySales ("All")
End Sub
Private Sub OKbutton_Click()
    CheckingForm.Hide
    CheckPaySales (SalesList.Value)
End Sub

Private Sub UserForm_Click()

End Sub
