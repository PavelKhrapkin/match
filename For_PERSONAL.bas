Attribute VB_Name = "For_PERSONAL"
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
'   26.7.2012 - match 2.0 - MoveToMatch � �������������� TOCmatch
'   17.8.2012 - ��������� ��������� - Loader'�� � ProcessEngine
'    8.9.2012 - ���� ������ ������� ��� ��������� ForPERSONAL.bas, ����� �� ������

    Option Explicit    ' Force explicit variable declaration
    
Sub MoveToMatch()
Attribute MoveToMatch.VB_Description = "20.7.12 MoveToMatch of Application match2.0"
Attribute MoveToMatch.VB_ProcData.VB_Invoke_Func = "�\n14"
'
' <*> MoveToMatch() - ����������� �������� ������ � ���� � ������ ��� ���������
'
' Keyboard Shortcut: Ctrl+�     -- Ctrl/� �����������, ����� �� ������������ Shift,
'                                  ��������������� ���������� �� Open
'
'Pavel Khrapkin 23-Dec-2011
' 8.2.2012 - ���������� ����� �����, ��������� ��� ���������
' 26.7.12 - match2.0 - ������������� ������ �� ���
' 1.8.12 - RepTOC.EOL ������ ������ EOL(TOC,DB_MATCH), bug fix
'          ����� ���� ���������, ���������� � ����������� ����������
' 11.8.12 - bug fix - ��������� ���� ������
' 18.8.12 - ������� ��������� ���� � MoveInMatch � match.xlsm
' 10.9.12 - bug fix - �� ��� ���� Path DBs
    
    Dim D As String
    D = "C:\work\Match\match2.0\DBs"
    Const F = "match.xlsm"
    Const DinC = "C:\match_environment.xlsx"
    Dim P As String
    
    Dim NewRep As String    ' ��� ����� � ����� �������
    NewRep = ActiveWorkbook.Name
    If NewRep = "Book1" Or NewRep = "PERSONAL.XLSB" Then
        MsgBox "������ ������ ������� �������� ��� PERSONAL.xlsb", , "<ERROR!>"
        End
    End If

OpenTry:
    P = D & "\" & F
    
'------- ���� match.xlsm ���� ����� ���������� � D?
    Dim W As Workbook
    On Error Resume Next
    Set W = Workbooks.Open(P, UpdateLinks:=False)
    On Error GoTo 0
    If W Is Nothing Then
'------- ���, match.xlsm �� ���. �� ���� ����� Workbooks?
        For Each W In Application.Workbooks
            If W.Name = F Then
                P = W.Path & "\" & F
                GoTo RunMatch
            End If
        Next W
'------ � ����� Workbooks ���. ��������� � ����� � �:\
        On Error Resume Next
        Set W = Workbooks.Open(DinC)
        P = W.Sheets(1).Cells(1, 2) & F
        W.Close
        Set W = Workbooks.Open(P, UpdateLinks:=False)
        On Error GoTo 0
        If W Is Nothing Then
Const Msg = "<!> MoveToMatch �� ������� ������� ���� match.xlsm'" _
    & vbCrLf & vbCrLf & "�������� ������� ��� �������, � �����" _
    & vbCrLf & "��� ��� ������� MoveToMatch (Ctrl/�)"
            If MsgBox(Msg, vbYesNo) = vbYes Then GoTo OpenTry
            End
        End If
    End If
RunMatch:
    Workbooks(NewRep).Activate
    Application.Run "'" & P & "'!MoveInMatch"
        
    End Sub
Sub TriggerOptionsFormulaStyle()
Attribute TriggerOptionsFormulaStyle.VB_ProcData.VB_Invoke_Func = "R\n14"
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
