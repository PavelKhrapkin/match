Attribute VB_Name = "match2_0"
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
    
''''    Public Const DownloadDir = "C:\Users\������������\Downloads\"
''''
''''' �������� ������ - �������. ����������� �� ������� � Match.xlsm ����� �������
'''''   !!! ����� ���������� �������������� ����� ���� ������ ����������    !!!
'''''   !!! ... �������� �� �������                                         !!!
''''    Public PaidSheet As String  ' ������ ����� 1� �� ��������� �� ������ �����,
''''                                    ' !!! �� ��� ���� ���������������!!!
''''    Public DogSheet As String   ' ������ ����� 1� �� ��������� �� ������ �����,
''''                                    ' !!! �� ��� ���� ���������������!!!
''''    Public Const DogHeader = "DogovorHeader" ' ������� ��� DL Dogovor_Insert
''''
'''''    Public Const PartnerCenter = "PartnerCenter"    ' ��� ����� ������ ��
''''                                '                  PartnerCenter.Autodesk.com
''''    Public Const PaidContract = "P_PaidContract" ' ������� ����- ������ �����
''''                                '                   .. ���������� ����������
''''    Public Const PaidNoContract = "P_PaidNoContract" ' ������ ����� ��������
''''                                '               ��� ���������� - ��������!!!
'''''    Public Const PaidUpdate = "P_Update"    ' ������� ���� - ������ �����
''''                                '               �������� ��� DL - ��������!!!
Sub MoveToMatch()
Attribute MoveToMatch.VB_Description = "8.2.2012 - ����������� �������� ������ �� ������ ���� MatchSF-1C.xlsb,  ������������� ��� �� ������ � ������ ������� �� ��� ������ "
Attribute MoveToMatch.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' <*> MoveToMatch() - ����������� �������� ������ � ���� � ������ ��� ���������
'
' Keyboard Shortcut: Ctrl+Shift+M
'
'Pavel Khrapkin 23-Dec-2011
' 8.2.2012 - ���������� ����� �����, ��������� ��� ���������
' 11.7.12 - match2.0 - ������������� ������, ������� ��� � ���� �� ������ ���� � ������ ���������

    Dim NewRep As String            ' ��� ����� � ����� �������
    Dim RepName As String           ' ��� ������ ������
    
    NewRep = ActiveWorkbook.Name
    Lines = EOL(1, Workbooks(NewRep))

    Set DB_MATCH = Workbooks.Open(F_MATCH, UpdateLinks:=False)
    
    Dim iDBs As Integer         '�������� TOCmatch - ���������� ��� ������
    iDBs = DB_MATCH.Sheets(1).Cells(4, TOC_PAR_1_COL)
    
'------ ������������� ������ NewRep �� ������� TOCmatch -------------
    Dim TOCline As Range        '= ������ TOC match
    With TOCline
        For Each TOCline In Range(Cells(5, 1), Cells(5 + iDBs, BIG)).Rows
            If IsRightStamp(TOCline, NewRep) Then GoTo RepNameCheck
        Next TOCline
        GoTo FatalNewRep
RepNameCheck:
        Dim FrTOC As Integer, ToTOC As Integer  '������ ������ RepName � TOC
        FrTOC = .Cells(1, TOC_PAR_2_COL)
        ToTOC = .Cells(1, TOC_PAR_3_COL)
        For Each TOCline In Range(Cells(FrTOC, 1), Cells(ToTOC, TOC_PAR_3_COL)).Rows
            If IsRightStamp(TOCline, NewRep) Then GoTo RepNameHandle
        Next TOCline
        GoTo FatalNewRep
RepNameHandle:
        
    End With
        
    If InStr(LCase$(Cells(Lines + 3, 1)), "salesforce.com") <> 0 _
            And Cells(Lines + SFresLines, 1) = SFstamp Then
            
        Set DB_SFDC = Workbooks.Open(F_SFDC, UpdateLinks:=False)
        Workbooks(NewRep).Sheets(1).Move Before:=DB_SFDC.Sheets(1)
        RepName = Cells(Lines + 2, 1)
        Lines = Lines - SFresLines
        
        Select Case RepName
            Case SFpayRepName:
                Application.Run ("SFDC.xlsm!Match1C_SF")    ' ����� SF �� ��������
            Case SFcontrRepName:
                Application.Run ("SFDC.xlsm!SFDreport")     ' ����� SFcontr - ��������
            Case SFaccRepName:
                Application.Run ("SFDC.xlsm!SFaccRep")      ' ����� SFacc - �����������
            Case SFcontactRepName:
                Application.Run ("SFDC.xlsm!SFcontactRep")  ' ����� SFcont �� ���������
            Case SFoppRepName:
                Application.Run ("SFDC.xlsm!SFoppRep")      ' ����� SFopp �� ��������
            Case SFadskRepName:
                Application.Run ("SFDC.xlsm!ADSKfromSFrep") ' ����� SF �� Autodesk
            Case SFpaRepName:
                Application.Run ("SFDC.xlsm!SF_PA_Rep")     ' ����� �� ������� �������� � ADSK
            Case Else:
                ErrMsg FATAL_ERR, "�� ��������� ����� Salesforce.com"
        End Select

        
        '** ������ 1� � Autodesk **
    ElseIf Cells(1, 1) = Stamp1Cpay1 And Cells(1, 2) = Stamp1Cpay2 Then
        Application.Run ("1C.xlsm!From1Cpayment")    ' ����� 1� �� ��������
    ElseIf Cells(1, 2) = Stamp1Cdog1 And Cells(1, 4) = Stamp1Cdog2 Then
        Application.Run ("1C.xlsm!From1Cdogovor")    ' ����� 1� �� ���������
    ElseIf Cells(1, 5) = Stamp1Cacc1 And Cells(1, 6) = Stamp1Cacc2 Then
        Application.Run ("1C.xlsm!From1Caccount")    ' ����� 1� �� ��������
'''    ElseIf Cells(1, 40) = StampADSKp1 And Cells(1, 42) = StampADSKp2 Then
'''        FrPartnerCenter
    Else: GoTo FatalNewRep
        End
    End If
        
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    Workbooks(NewRep).Close

'------------- match TOC write -------------------------
    
'''    With DB_MATCH
'''            If TOCline.Cells(1, TOC_REPNAME_COL) = RepName Then
'''                With TOCline
'''                    .Cells(1, TOC_LOAD_COL) = Now
'''                    .Cells(1, TOC_HANDLE_COL) = ""
'''                    .Cells(1, TOC_EOL_COL) = Lines
'''                End With
'''                Exit For
'''            End If
'''        .Cells(1, 1) = Now
'''        .Save
'''    End With
    Exit Sub
FatalNewRep:
    ErrMsg FATAL_ERR, "������� ����� '" & NewRep & "' �� ���������"
End Sub
Function IsRightStamp(TOCline, NewRep) As Boolean
'
' - IsRightStamp(TOCline) - �������� ������������ ������ � NewRep �� ������ TOCline
' 12.7.2012

    Dim NewRepStamp As String       ' ����� ������ ������
    
    Dim Stamp As String         '= ������ - �����
    Dim StampType As String     '��� ������: ������ (=) ��� ���������
    Dim Stamp_R As Integer      '����� ������, ��� �����
    Dim Stamp_C As Integer      '����� �������, ��� �����
    Dim ParCheck As Integer     '�������� TOCmatch - ������ �������������� �������� ������
    
    IsRightStamp = False
        
    With TOCline
        Do
            Stamp = .Cells(1, TOC_STAMP_COL)
            If Stamp = "" Then Exit Function        ' ����������� ����� - �� �������!
            StampType = .Cells(1, TOC_STAMP_TYPE_COL)
            Stamp_R = .Cells(1, TOC_STAMP_R_COL)
            If .Cells(1, TOC_EOL_COL) = "EOL" Then Stamp_R = Stamp_R + Lines
            Stamp_C = .Cells(1, TOC_STAMP_C_COL)
            NewRepStamp = Workbooks(NewRep).Sheets(1).Cells(Stamp_R, Stamp_C)
            
            If StampType = "=" And NewRepStamp <> Stamp Then
                Exit Function
            ElseIf StampType = "I" And InStr(LCase$(NewRepStamp), LCase$(Stamp)) = 0 Then
                Exit Function
            Else: If StampType <> "=" And StampType <> "I" Then _
                ErrMsg FATAL_ERR, "���� � ��������� TOCmatch: ��� ������ =" & StampType
            End If
        
            ParCheck = .Cells(1, TOC_PAR_1_COL)
            If IsNumeric(ParCheck) And ParCheck > 0 Then
                Set TOCline = Range(Cells(ParCheck, 1), Cells(ParCheck, BIG))
            End If
        Loop While ParCheck <> 0
    End With

    IsRightStamp = True

End Function
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
