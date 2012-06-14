Attribute VB_Name = "Match1C"
'---------------------------------------------------------------------------
' ������� ��� ������ � ������ ������� �� 1� � Salesforce Match SF-1C.xlms
'
' * MoveToMatch    - ������� ����� �� ������ ����� Match1SF    (Ctrl/Shift/M)
' * TriggerOptionsFormulaStyle  - ������������ ���� A1/R1C1    (Ctrl/Shift/R)
'
' �.�.������� 2.1.2012
'   28.1.2012 - ������ �� �������������� ���� � ������� ������
'    5.2.2012 - � MoveToMatch - ������������� �������� ������ �� ������
'   16.5.2012 - �������� ����� SF_PA
'    2.6.2012 - TriggerOptionsFormulaStyle A1/R1C1

    Option Explicit    ' Force explicit variable declaration
    
    Public Const DownloadDir = "C:\Users\������������\Downloads\"

' �������� ������ - �������. ����������� �� ������� � Match.xlsm ����� �������
'   !!! ����� ���������� �������������� ����� ���� ������ ����������    !!!
'   !!! ... �������� �� �������                                         !!!
    Public PaidSheet As String  ' ������ ����� 1� �� ��������� �� ������ �����,
                                    ' !!! �� ��� ���� ���������������!!!
    Public DogSheet As String   ' ������ ����� 1� �� ��������� �� ������ �����,
                                    ' !!! �� ��� ���� ���������������!!!
    Public Const DogHeader = "DogovorHeader" ' ������� ��� DL Dogovor_Insert

'    Public Const PartnerCenter = "PartnerCenter"    ' ��� ����� ������ ��
                                '                  PartnerCenter.Autodesk.com
    Public Const PaidContract = "P_PaidContract" ' ������� ����- ������ �����
                                '                   .. ���������� ����������
    Public Const PaidNoContract = "P_PaidNoContract" ' ������ ����� ��������
                                '               ��� ���������� - ��������!!!
'    Public Const PaidUpdate = "P_Update"    ' ������� ���� - ������ �����
                                '               �������� ��� DL - ��������!!!
    Public Lines As Integer     ' ���������� ����� ��������/������ ������
    Public LinesOld As Integer  ' ���������� ����� ������� ������
    Public AllCol As Integer    ' ���������� ������� � ������� ������
    Public Doing As String      ' ������ � Application.StatusBar - ��� ������ ������
Sub MoveToMatch()
Attribute MoveToMatch.VB_Description = "8.2.2012 - ����������� �������� ������ �� ������ ���� MatchSF-1C.xlsb,  ������������� ��� �� ������ � ������ ������� �� ��� ������ "
Attribute MoveToMatch.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' 1. Macro to match SF and 1C Reports Pavel Khrapkin 23-Dec-2011
'
' Keyboard Shortcut: Ctrl+Shift+M
'
'       8.2.2012 - ���������� ����� �����, ��������� ��� ���������

    Dim SFrepName As String
    Const ErMsg = "�� ������ ����� �������������� ����� �����"
    
    Sheets(1).Move Before:=Workbooks("Match SF-1C.xlsm").Sheets(1)
    
    Lines = EOL(1)
        
    If Cells(Lines, 1) = SFstamp Then                               '** ����� SF **
        Select Case Cells(Lines - 4, 1)
            Case SFpayRepName: Call Match1C_SF  ' ����� SF �� ��������
            Case SFcontrRepName: Call SFDreport ' ����� SFcontr - ��������
            Case SFaccRepName: Call SFaccRep    ' ����� SFacc - �����������
            Case SFoppRepName: Call SFoppRep    ' ����� SFopp �� ��������
            Case SFadskRepName: ADSKfromSFrep   ' ����� SF �� Autodesk
            Case SFpaRepName: SF_PA_Rep         ' ����� ������� �������� � ADSK
            Case Else: MsgBox ErMsg & " Salesforce.", , "���������!"
        End Select                                              '** ������ 1� � Autodesk **
    ElseIf Cells(1, 1) = Stamp1Cpay1 And Cells(1, 2) = Stamp1Cpay2 Then From1Cpayment
    ElseIf Cells(1, 2) = Stamp1Cdog1 And Cells(1, 4) = Stamp1Cdog2 Then From1Cdogovor
    ElseIf Cells(1, 5) = Stamp1Cacc1 And Cells(1, 6) = Stamp1Cacc2 Then From1Caccount
    ElseIf Cells(1, 40) = StampADSKp1 And Cells(1, 42) = StampADSKp2 Then FrPartnerCenter
    Else: MsgBox ErMsg, , "���������!"
    End If
    ActiveWorkbook.Save
End Sub
Sub TriggerOptionsFormulaStyle()
'
' * Trigger Options-Formula Style A1/R1C1
'
' CTRL+Shift+R
'
' 2.6.12
    If Application.ReferenceStyle = xlR1C1 Then
        Application.ReferenceStyle = xlA1
    Else
        Application.ReferenceStyle = xlR1C1
    End If
End Sub
