VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainControlPanel 
   Caption         =   "Основная управляющая панель Match (Ctrl/Shift/Q)"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5775
   OleObjectBlob   =   "MainControlPanel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-------------------------------------------------------
' [**] MainContrPanel - основная панел запуска процедур Match
'   1.3.2012
' 23.5.12 - кнопка A_DicBulid
' 28.5.12 - кнопка Subscriptions ADSK
'  8.6.12 - кнопка Отчеты ADSK

Option Explicit

Private Sub ADSKreportsButton_Click()
    MainControlPanel.Hide   ' загрузка и проход по отчету из ADSK.xlsx
    ADSK_TOC_FormOutput
End Sub

Private Sub SubscriptionsADSKbutton_Click()
    MainControlPanel.Hide   ' проход по листу Subscriptions из ADSK PartnerCenter
    SubscriptionsADSKpass
End Sub

Private Sub ContractPassButton_Click()
    MainControlPanel.Hide   ' проход по листу Договоров 1С
    ContrPass
End Sub

Private Sub P_HandlingButton_Click()
    MainControlPanel.Hide
    PaidHandling        ' проход по листу Платежей (1)
End Sub
Private Sub StockPassButton_Click()
' [*] проход по Складской Книге
'   18.5.12
    MainControlPanel.Hide
    StockHandling
End Sub
Private Sub A_DicBuildButton_Click()
    MainControlPanel.Hide
    SFaccDicBuild
End Sub

Private Sub CheckingButton_Click()
    MainControlPanel.Hide
'    CheckingForm.Show
    CheckFofmOutput     ' вывод CheckingForm с формированием списка Продавцов
End Sub
Private Sub Debug3passForm_Click()
    MainControlPanel.Hide
    Debug3pass.Show
    End
End Sub
Private Sub NewDogButton_Click()
    MainControlPanel.Hide
    NewContractDL       ' создание новых Договоров по листу Договоров (3)
End Sub

Private Sub UserForm_Click()

End Sub
