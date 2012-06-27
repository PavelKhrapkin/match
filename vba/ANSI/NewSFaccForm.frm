VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewSFaccForm 
   Caption         =   "Создание организации SF"
   ClientHeight    =   10845
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
' 27.06.12




Option Explicit
Dim FaxFromTel As Boolean
Sub setFaxfromTel(par)
    FaxFromTel = par
End Sub
Private Sub CancelButton_Click()
    Me.result.value = "cancel"
    Me.Hide
End Sub
Private Sub BackButton_Click()
    Me.result.value = "back"
    Me.Hide
End Sub

Private Sub ExitButton_Click()
    Me.result.value = "exit"
    Me.Hide
End Sub
Private Sub AccSaveForSF_Click()
    
    If Not checkDigits(Me.INN, 10) Then
        MsgBox "ИНН неверен. Он должен состоять из 10 цифр."
    ElseIf checkAddr() Then
        Me.result.value = "save"
        Me.Hide
    End If
End Sub

Private Sub City_Change()
'    MsgBox "city change"
End Sub

Private Sub Label19_Click()

End Sub

Private Sub phone_Change()
    If FaxFromTel Then Me.fax = telToFax(Me.phone)
End Sub
Private Function checkAddr() As Boolean
    Dim i As Long
    checkAddr = checkIndex(Me.Index)
    If Not checkAddr Then
        MsgBox "индекс должен состоять из 6 цифр"
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
            MsgBox "индекс факт. адреса должен состоять из 6 цифр"
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
' проверить корректность почтового индекса / ИНН : должно быть lng символов и все - цифры.
    Dim i As Long
    checkDigits = False
    If Len(Index) <> lng Then GoTo exitFunction
    For i = 1 To lng
        If Not IsNumeric(Mid(Index, i, 1)) Then GoTo exitFunction
    Next
    checkDigits = True
exitFunction:
End Function

Sub setPostAddr(addr As PostAddr)
    Me.Area = addr.State
    Me.City = addr.City
    Me.Street = addr.Street
    Me.Index = addr.PostIndex
    Me.Country = addr.Country
End Sub
Sub setDelAddr(addr As PostAddr)
    Me.AreaD = addr.State
    Me.CityD = addr.City
    Me.StreetD = addr.Street
    Me.IndexD = addr.PostIndex
    Me.CountryD = addr.Country
End Sub
Private Sub Label7_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub title1C_Click()

End Sub

Private Sub UserForm_Click()

End Sub
