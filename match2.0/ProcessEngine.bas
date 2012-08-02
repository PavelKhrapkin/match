Attribute VB_Name = "ProcessEngine"
'----------------------------------------------------------------------------------------------------------
' ProcessEngine - ��������� ���������. ��������� � ���� ��������� �������� � ����� Process ����� match.xlsm
'                   * �������� Loader'� ����������� ��� �������� � DBs ������ ���������.
'                   * �������� ������ ��������� ���������� � ��������� ��������� ��� ��������,
'                     ��������� � �������������� ����� ���������
'                   * �������� Handler'� � Reporter'� ������������ �� ���� ����� ������� ������
'                   * ��� �������� ���������������. ���� ��� �������� ��� �������� - �� ������������.
'
'   2.8.12 �.�.�������
'
' - ProcStart(Proc) - ������ �������� Proc �� ������� Process � match.xlsm

Option Explicit

Sub ProcStart(Proc)
'
' - ProcStart(Proc) - ������ �������� Proc �� ������� Process � match.xlsm
'   2.8.12

    Dim i As Integer
    Dim Step As String          ' ��� ��������
    Dim PrevStep As String      ' ���������� ������������ ���
    Dim ProcName As String      ' ��� ��������
    
    GetRep Process
    With DB_MATCH.Sheets(Process)
        For i = 6 To RepTOC.EOL
            Step = .Cells(i, PROC_STEP_COL)
            ProcName = .Cells(i, PROC_NAME_COL)
            If Step = PROC_START And ProcName = Proc Then GoTo MyProc
        Next i
        ErrMsg FATAL_ERR, "�� ������ ������� " & Proc
        End
MyProc:
        Dim StepRow As Range
        Do While Step <> PROC_END
            i = i + 1
            Step = .Cells(i, PROC_STEP_COL)
            
'-- �������� -?-IsDone -- ��������, ��� ������������ �������� ���������
'                      -- ���� ��� �� �������� - ��� ����������� �����
            If Step = PROC_IFISDONE Then
                ProcName = .Cells(i, PROC_NAME_COL)
                Step = .Cells(i, PROC_NAME_COL + 1)
                If Not IsDone(ProcName, Step) Then ProcStart ProcName
            End If
            
            If .Cells(i, PROC_STEPDONE_COL) <> "1" Then
                PrevStep = .Cells(i, PROC_PREVSTEP_COL)
'-- ���������, �������� �� ���������� ��������� ��� �� PrevStep
                If PrevStep <> REP_LOADED Then
                    GetRep .Cells(i, PROC_REP1_COL)
                    If RepTOC.Made <> REP_LOADED Then GoTo Err
                End If
                If Not IsDone(Proc, PrevStep) Then GoTo Err

'*************************************
                Exec Step, i        '*  ��������� ���
'*************************************
            
            End If
        Loop
        
    End With
    MS "<*> ������� " & Proc & " ��������!"
    Exit Sub
Err:
    ErrMsg FATAL_ERR, "�������� ������������������ ����� ��������� " & ProcName
    End
End Sub
Function IsDone(Proc, Step) As Boolean
'
' - IsDone(Proc, Step) - ��������, ��� ��� Step �������� Proc ��������
'   2.8.12

    Dim P As TOCmatch
    Dim i As Integer
    Dim ProcName As String
    Dim StepName As String
    
    P = GetRep(Process)
    
    With DB_MATCH.Sheets(Process)
        For i = 6 To P.EOL
            ProcName = .Cells(i, PROC_NAME_COL)
            StepName = .Cells(i, PROC_STEP_COL)
            If StepName = PROC_START And ProcName = Proc Then GoTo MyProc
        Next i
        ErrMsg FATAL_ERR, "IsDone: ��������� � ��������������� �������� " & Proc
        Stop
        End
MyProc: Do While StepName <> PROC_END
            i = i + 1
            StepName = .Cells(i, PROC_STEP_COL)
            If StepName = Step Then
                IsDone = False
                If .Cells(i, PROC_STEPDONE_COL) <> "1" Then IsDone = True
                Exit Function
            End If
        Loop
        ErrMsg FATAL_ERR, "IsDone: ��������� � ��������������� ���� " & Step _
            & " �������� " & Proc
    End With
End Function
Sub Exec(Step, iProc)
'
' - Exec(Step, iProc) - ����� ���� Step �� ������ iProc ������� ���������
'   2.8.12
       
    Dim Code As String
    Dim File As String
       
    With DB_MATCH.Sheets(Process)
        Code = Step
        File = .Cells(iProc, PROC_STEPFILE_COL)
        If File <> "" Then Code = "'" & DirDBs & File & "'!" & Step
        
        If .Cells(iProc, PROC_PAR1_COL + 4) <> "" Then
            Application.Run Code, _
                .Cells(iProc, PROC_PAR1_COL), _
                .Cells(iProc, PROC_PAR1_COL + 1), _
                .Cells(iProc, PROC_PAR1_COL + 2), _
                .Cells(iProc, PROC_PAR1_COL + 3), _
                .Cells(iProc, PROC_PAR1_COL + 4)
        ElseIf .Cells(iProc, PROC_PAR1_COL + 3) <> "" Then
            Application.Run Code, _
                .Cells(iProc, PROC_PAR1_COL), _
                .Cells(iProc, PROC_PAR1_COL + 1), _
                .Cells(iProc, PROC_PAR1_COL + 2), _
                .Cells(iProc, PROC_PAR1_COL + 3)
        ElseIf .Cells(iProc, PROC_PAR1_COL + 2) <> "" Then
            Application.Run Code, _
                .Cells(iProc, PROC_PAR1_COL), _
                .Cells(iProc, PROC_PAR1_COL + 1), _
                .Cells(iProc, PROC_PAR1_COL + 2)
        ElseIf .Cells(iProc, PROC_PAR1_COL + 1) <> "" Then
            Application.Run Code, _
                .Cells(iProc, PROC_PAR1_COL), _
                .Cells(iProc, PROC_PAR1_COL + 1)
        ElseIf .Cells(iProc, PROC_PAR1_COL) <> "" Then
            Application.Run Code, _
                .Cells(iProc, PROC_PAR1_COL)
        Else
            Application.Run Code
        End If
        .Cells(iProc, PROC_STEPDONE_COL) = "1"  ' Done = "1" - ��� ��������
        .Cells(iProc, PROC_TIME_COL) = Now
        .Cells(1, 1) = Now
    End With
End Sub
Sub testRunProc()
    RunProc "REP_SF_LOAD"
End Sub
Sub RunProc(Proc)
'
' - RunProc(Proc)   - ������ �������� Proc
'   31.7.12
'''    GetRep Process
    Application.Run "'" & DirDBs & F_MATCH & "'!ProcStart", Proc
End Sub
