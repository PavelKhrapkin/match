VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "Ход выполнения процедуры Match"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4380
   OleObjectBlob   =   "ProgressForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ProgressFrame_Click()
    StopSub
End Sub

Private Sub StopButton_Click()
    LogWr "STOP в Progress"
    ExRespond = False
End Sub

Private Sub UserForm_Click()
    StopSub
End Sub
