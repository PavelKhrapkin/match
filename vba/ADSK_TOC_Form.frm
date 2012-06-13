VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ADSK_TOC_Form 
   Caption         =   "Выбор отчета по Содержанию ADSK.xlsx"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "ADSK_TOC_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ADSK_TOC_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub OKbutton_Click()
    ADSK_TOC_Form.Hide
    ADSK_ReportHandle TOClist.value
End Sub
Private Sub UserForm_Click()

End Sub
