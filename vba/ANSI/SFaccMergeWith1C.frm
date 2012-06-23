VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SFaccMergeWith1C 
   Caption         =   "Модифицировать связываемое SF предприятие"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20955
   OleObjectBlob   =   "SFaccMergeWith1C.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SFaccMergeWith1C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim INN As String
Dim tel As String
Dim adr As PostAddr, delAddr As PostAddr

Function getInn()
    getInn = Me.innSF
    If Me.chkInn2 Then getInn = Me.inn1C
End Function
Function getTel()
    getTel = Me.telSF
    If Me.chkInn2 Then getTel = Me.tel1C
End Function
Function setTel(pSF, p1C)
    tel = pSF                       ' запомнить первоначальное значение
    Me.telSF = pSF
    Me.tel1C = p1C
    Me.chkTel1 = False              ' привести chekboxes в исходное состояние
    Me.chkTel2 = False
End Function
Function setInn(pSF As String, p1C As String)
    INN = ""
    If pSF <> "" Then INN = split(pSF, "/")(0)       ' запомнить первоначальное значение
    Me.innSF = pSF
    Me.inn1C = split(p1C, "/")(0)
    Me.chkInn1 = False              ' привести chekboxes в исходное состояние
    Me.chkInn2 = False
End Function
Function setAddr(adSF As PostAddr, ad1C As PostAddr, delAddrSF As PostAddr, factAddr1C As PostAddr)

    ' основной адрес
    adr = adSF
    Me.CitySF = adSF.City
    Me.AreaSF = adSF.State
    Me.StreetSF = adSF.Street
    Me.IndexSF = adSF.PostIndex
    Me.CountrySF = adSF.Country
    Me.City1C = ad1C.City
    Me.Area1C = ad1C.State
    Me.Street1C = ad1C.Street
    Me.Index1C = ad1C.PostIndex
    Me.Country1C = ad1C.Country
    Me.chkAdr1 = False              ' привести chekboxes в исходное состояние
    Me.chkAdr2 = False
    
    ' адрес доставки SF / фактический 1С
    delAddr = delAddrSF
    Me.DelCitySF = delAddrSF.City
    Me.DelAreaSF = delAddrSF.State
    Me.DelStreetSF = delAddrSF.Street
    Me.DelIndexSF = delAddrSF.PostIndex
    Me.DelCountrySF = delAddrSF.Country
    Me.FactCity1C = factAddr1C.City
    Me.FactArea1C = factAddr1C.State
    Me.FactStreet1C = factAddr1C.Street
    Me.FactIndex1C = factAddr1C.PostIndex
    Me.FactCountry1C = factAddr1C.Country
    Me.chkDelAddr1 = False          ' привести chekboxes в исходное состояние
    Me.chkDelAddr2 = False
End Function
Sub chkInn1_Change()
    If Me.chkInn1 Then Me.innSF = INN
End Sub
Sub chkInn2_Change()
    If Me.chkInn2 Then Me.innSF = Me.inn1C
End Sub
Sub chkTel1_Change()
    If Me.chkTel1 Then Me.telSF = tel
End Sub
Sub chkTel2_Change()
    If Me.chkTel2 Then Me.telSF = Me.tel1C
End Sub
Sub chkAdr1_Change()
    If Me.chkAdr1 Then
        Me.CitySF = adr.City
        Me.StreetSF = adr.Street
        Me.AreaSF = adr.State
        Me.IndexSF = adr.PostIndex
        Me.CountrySF = adr.Country
    End If
End Sub
Sub chkAdr2_Change()
    If Me.chkAdr2 Then
        Me.CitySF = Me.City1C
        Me.StreetSF = Me.Street1C
        Me.AreaSF = Me.Area1C
        Me.IndexSF = Me.Index1C
        Me.CountrySF = Me.Country1C
    End If
End Sub
Sub chkDelAddr1_Change()
    If Me.chkDelAddr1 Then
        Me.DelCitySF = delAddr.City
        Me.DelStreetSF = delAddr.Street
        Me.DelAreaSF = delAddr.State
        Me.DelIndexSF = delAddr.PostIndex
        Me.DelCountrySF = delAddr.Country
    End If
End Sub
Sub chkDelAddr2_Change()
    If Me.chkDelAddr2 Then
        Me.DelCitySF = Me.FactCity1C
        Me.DelStreetSF = Me.FactStreet1C
        Me.DelAreaSF = Me.FactArea1C
        Me.DelIndexSF = Me.FactIndex1C
        Me.DelCountrySF = Me.FactCountry1C
    End If
End Sub

Private Sub inn1C_Click()

End Sub

Private Sub telSF_Change()
'    MsgBox "phone change"
    Me.faxSF = telToFax(Me.telSF)
End Sub
'Sub innSF_Change()
'    innRes = Me.innSF              ' кнопки 'reset' удалены 16.06.12
'End Sub
'Private Sub resetInn_Click()
'    Me.chkInn2 = False
'    Me.innSF = inn
'End Sub
'Private Sub resetTel_Click()
'    Me.chkTel1 = False
'    Me.chkTel2 = False
'    Me.telSF = inn
'End Sub

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
    Me.result = "save"
    INN = Me.inn1C
    Me.Hide
End Sub

Private Sub Label42_Click()

End Sub

Private Sub UserForm_Click()

End Sub
