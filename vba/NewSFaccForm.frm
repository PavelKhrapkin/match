VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewSFaccForm 
   Caption         =   "Организация - Клиент 1С"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10710
   OleObjectBlob   =   "NewSFaccForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewSFaccForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub CancelButton_Click()
    Me.result.value = "cancel"
    Me.Hide
End Sub
Private Sub AccSaveForSF_Click()
    
    If Not checkDigits(Me.INN, 10) Then
        MsgBox "неверный ИНН"
    ElseIf checkAddr() Then
        Me.result.value = "save"
        Me.Hide
    End If
End Sub

Private Sub City_Change()
'    MsgBox "city change"
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label6_Click()

End Sub
Private Function checkAddr() As Boolean
    Dim i As Long
    checkAddr = checkIndex(Me.Index)
    If Not checkAddr Then
        MsgBox "неверный индекс"
    End If
    If Trim(Me.Street) = "" Then
        MsgBox "не введено поле 'улица'"
        checkAddr = False
    End If
    If Trim(Me.City) = "" Then
        MsgBox "не введено поле 'город'"
        checkAddr = False
    End If
    If Me.IndexD <> "" Or Me.StreetD <> "" Or Me.CityD <> "" Then
        If Not checkIndex(Me.IndexD) Then
            MsgBox "неверный индекс факт. адреса"
        End If
        If Trim(Me.StreetD) = "" Then
            MsgBox "не введено поле 'улица' факт. адреса"
            checkAddr = False
        End If
        If Trim(Me.CityD) = "" Then
            MsgBox "не введено поле 'город' факт. адреса"
            checkAddr = False
        End If
        If Trim(Me.CountryD) = "" Then
            Me.CountryD = "Россия"
        End If
    End If
End Function
Function checkIndex(Index) As Boolean
    checkIndex = checkDigits(Index, 6)
End Function
Function checkDigits(Index, ByVal lng As Long) As Boolean
' проверить корректность почтового индекса: должно быть 6 символов и все - цифры.
    Dim i As Long
    checkDigits = False
    If Len(Index) <> lng Then GoTo exitFunction
    For i = 1 To lng
        If Not IsNumeric(Mid(Index, i, 1)) Then GoTo exitFunction
    Next
    checkDigits = True
exitFunction:
End Function

Private Sub Label7_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Click()

End Sub
