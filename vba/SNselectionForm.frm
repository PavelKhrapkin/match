VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SNselectionForm 
   Caption         =   "Выбор SN"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1890
   OleObjectBlob   =   "SNselectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SNselectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
' Работа с SNselectionForm - выбор SN для процедуры 3PASS
'   8.2.2012

Option Explicit
Private Sub OKbutton_Click()
'
' [OK] - выбор группы SN из ADSKfrSF и заполнением списком SN колонки А листа 3PASS
' 8.2.2012

    Const SN = 4        ' номер колонки "SN продукта Autodesk" в ADSKfrSF
    Const StatusSN = 7  ' номер колонки "Статус SN" в ADSKfrSF
    Const SeatsSN = 6   ' номер колонки "Seats" (число мест) в ADSKfrSF
    Dim Status As String    ' Статус текущего SN
    Dim Seats               ' число посадочных мест для текущего SN
    Dim FrN, ToN, i
' заполняем колонку А листа 3PASS выбранными серийнами номерами
    FrN = 5
    ToN = FrN
    For i = 2 To EOL(ADSKfrSF)
        If Sheets(ADSKfrSF).Cells(i, SN) = "" Then i = i + 1 ' пустые SN игнорируем
        Status = Sheets(ADSKfrSF).Cells(i, StatusSN)
        Seats = Sheets(ADSKfrSF).Cells(i, SeatsSN)
        If SNregistered And Status = "Registered" Or _
           SNunregistered And Status = "Untegistered" Or _
           SNupgraded And Status = "Upgraded" Or _
           SNnotUpgradeable And Status = "Not Upgradeable" Or _
           SNnew And Seats > 777 _
           Then
            Sheets(A3PASS).Cells(ToN, 1) = Sheets(ADSKfrSF).Cells(i, SN) & "+"
            ToN = ToN + 1
        End If
    Next i
    
    Sheets(A3PASS).Select
    If FrN = ToN Then    ' ни одного SN не выбрано - выбираем еще раз
        MsgBox "Выбери группу SN!"
        Exit Sub
    End If
    Cells(2, 1) = FrN
    Cells(3, 1) = ToN - 1
    SNselectionForm.Hide
End Sub
Private Sub CancelButton_Click()
'
' [Cancel] - отмена работы с 3PASS
'   7.2.2012

    End
End Sub

Private Sub UserForm_Click()

End Sub
