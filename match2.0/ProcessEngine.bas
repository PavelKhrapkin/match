Attribute VB_Name = "ProcessEngine"
'---------------------------------------------------------------------------------------
' ProcessEngine - ��������� ���������.
'         * ��������� � ���� ��������� �������� � Process ����� match.xlsm
'         * �������� Loader'� ����������� ��� �������� � DBs ������ ���������.
'         * �������� ������ ��������� ���������� � ��������� ��������� ���
'           ��������,��������� � �������������� ������� ���������
'         * �������� Handler'� � Reporter'� ������������ �� ���� ����� ������� ������
'         * ��� �������� ���������������. ���� ��� ��� �������� - �� ������������.
'         * ����� ����������� ���� ����������� ���� Done �� ���� PrevStep.
'           PrevStep ����� ����� ��� <������ �������> / <���>.
'
'   4.8.12 �.�.�������
'
' - ProcStart(Proc)     - ������ �������� Proc �� ������� Process � match.xlsm
' - IsDone(Proc, Step)  - ��������, ��� ��� Step �������� Proc ��������
' - Exec(Step, iProc)   - ����� ���� Step �� ������ iProc ������� ���������
' - ToStep(Proc,[Step]) - ���������� ����� ������ ������� ���������

Option Explicit

Sub ProcStart(Proc As String)
'
' - ProcStart(Proc) - ������ �������� Proc �� ������� Process � match.xlsm
'   2.8.12

    Dim Step As String, PrevStep As String
    Dim i As Integer
    
    Proc = Trim(Proc)
    
    i = ToStep(Proc)
    With DB_MATCH.Sheets(Process)
        Do While Step <> PROC_END
            i = i + 1
            Step = .Cells(i, PROC_STEP_COL)
            
            If .Cells(i, PROC_STEPDONE_COL) <> "1" Then
                PrevStep = .Cells(i, PROC_PREVSTEP_COL)
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
    ErrMsg FATAL_ERR, "�������� ������������������ ����� ��������� " & Proc
    End
End Sub
Function IsDone(ByVal Proc As String, ByVal Step As String) As Boolean
'
' - IsDone(Proc, Step) - ��������, ��� ��� Step �������� Proc ��������
'   2.8.12

    Dim i As Integer
    Dim S() As String   '=����� ���������� PrevStep, ����������� ","
    Dim x() As String   '=������ ����� ����� ���� ���� <Proc>/<Step>
    Dim Rep As String, Done As String
    
    Proc = Trim(Proc): Step = Trim(Step)
    
    If InStr(Step, ",") <> 0 Then
        S = split(Trim(Step), ",")
        For i = LBound(S) To UBound(S)
            If InStr(S(i), "/") <> 0 Then
                x = split(S(i), "/")
                If Not IsDone(x(0), x(1)) Then ProcStart x(0)
            Else
                If Not IsDone(Proc, S(i)) Then ProcStart Proc
            End If
        Next i
    End If
        
    If Step = REP_LOADED Then
        i = ToStep(Proc)
        Rep = DB_MATCH.Sheets(Process).Cells(i, PROC_REP1_COL)
        GetRep Rep
        If RepTOC.Made <> REP_LOADED Then
            Dim Msg As String
            Msg = "IsDone: �� 'Loaded' ���� ��� �������� " _
                & Proc & " �� ���� " & Step & vbCrLf & vbCrLf _
                & "������ ���� ������� ������ ?"
            If MsgBox(Msg, vbYesNo) <> vbYes Then
                IsDone = False
                Exit Function
            Else
                ProcStart Proc
            End If
        Else
            IsDone = True
            Exit Function
        End If
    End If
    
    i = ToStep(Proc, Step)
    Done = DB_MATCH.Sheets(Process).Cells(i, PROC_STEPDONE_COL)
    
    IsDone = True
    If Done = "1" Then Exit Function
    IsDone = False
End Function
Sub Exec(Step, iProc)
'
' - Exec(Step, iProc) - ����� ���� Step �� ������ iProc ������� ���������
'   4.8.12
       
    Dim Code As String
    Dim File As String
    Dim R As TOCmatch       '= �������������� �������� - �����
          
    With DB_MATCH.Sheets(Process)
        R = GetRep(.Cells(iProc, PROC_REP1_COL))
        Workbooks(R.RepFile).Sheets(R.SheetN).Select    'Select ����, � ������� ����� ��������
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
'-- ������ ������� � ���� � TOCmatch � � ������� ���������
        .Cells(iProc, PROC_STEPDONE_COL) = "1"  ' Done = "1" - ��� ��������
        .Cells(iProc, PROC_TIME_COL) = Now
        .Cells(1, 1) = Now
        R.Made = Step
        R.Dat = Now
        RepTOC = R
        WrTOC
    End With
End Sub
Function ToStep(Proc, Optional Step As String = "") As Integer
'
' - ToStep(Proc, [Step]) - ���������� ����� ������ ������� ���������
'   2.8.12
    
    Dim P As TOCmatch           '������ ������� ��������� � ���� TOCmatch
    Dim StepName As String      '=��� �������� ����
    Dim ProcName As String      '=��� �������� ��������
    Dim i As Integer
    
    P = GetRep(Process)
    
    With DB_MATCH.Sheets(Process)
        For i = 6 To RepTOC.EOL
            ProcName = .Cells(i, PROC_NAME_COL)
            StepName = .Cells(i, PROC_STEP_COL)
            If StepName = PROC_START And ProcName = Proc Then GoTo MyProc
        Next i
        ErrMsg FATAL_ERR, "ToProc: �� ������ ������� " & Proc
        Stop
        End

MyProc: ToStep = i
        If Step = "" Then Exit Function
        Do While StepName <> PROC_END
            i = i + 1
            StepName = .Cells(i, PROC_STEP_COL)
            If StepName = Step Then
                ToStep = i
                Exit Function
            End If
        Loop
    End With
    ErrMsg FATAL_ERR, "ToStep: ��������� � ��������������� ���� " & Step _
        & " �������� " & Proc
End Function
Sub testRunProc()   'Ctrl/W
Attribute testRunProc.VB_ProcData.VB_Invoke_Func = "W\n14"
    RunProc "REP_SF_LOAD"
End Sub
Sub RunProc(Proc)
'
' - RunProc(Proc)   - ������ �������� Proc
'   31.7.12
'''    GetRep Process
    Application.Run "'" & DirDBs & F_MATCH & "'!ProcStart", Proc
End Sub
