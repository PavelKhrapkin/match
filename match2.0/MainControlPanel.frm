VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainControlPanel 
   Caption         =   "�������� ����������� ������ Match (Ctrl/Shift/Q)"
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
' [**] MainContrPanel - �������� ����� ������� �������� Match
'   1.3.2012
' 23.5.12 - ������ A_DicBulid
' 28.5.12 - ������ Subscriptions ADSK
'  8.6.12 - ������ ������ ADSK

Option Explicit

Private Sub ADSKreportsButton_Click()
    MainControlPanel.Hide   ' �������� � ������ �� ������ �� ADSK.xlsx
    ADSK_TOC_FormOutput
End Sub

Private Sub SubscriptionsADSKbutton_Click()
    MainControlPanel.Hide   ' ������ �� ����� Subscriptions �� ADSK PartnerCenter
    SubscriptionsADSKpass
End Sub

Private Sub ContractPassButton_Click()
    MainControlPanel.Hide   ' ������ �� ����� ��������� 1�
    ContrPass
End Sub

Private Sub P_HandlingButton_Click()
    MainControlPanel.Hide
    PaidHandling        ' ������ �� ����� �������� (1)
End Sub
Private Sub StockPassButton_Click()
' [*] ������ �� ��������� �����
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
    CheckFofmOutput     ' ����� CheckingForm � ������������� ������ ���������
End Sub
Private Sub Debug3passForm_Click()
    MainControlPanel.Hide
    Debug3pass.Show
    End
End Sub
Private Sub NewDogButton_Click()
    MainControlPanel.Hide
    NewContractDL       ' �������� ����� ��������� �� ����� ��������� (3)
End Sub

Private Sub UserForm_Click()

End Sub
