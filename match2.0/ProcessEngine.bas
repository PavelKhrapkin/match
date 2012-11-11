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
' 11.11.12 �.�.�������, �.����
'
' - ProcStart(Proc)     - ������ �������� Proc �� ������� Process � match.xlsm
' - IsDone(Proc, Step)  - ��������, ��� ��� Step �������� Proc ��� ��������
' - Exec(Step, iProc)   - ����� ���� Step �� ������ iProc ������� ���������
' - ToStep(Proc,[Step]) - ���������� ����� ������ ������� ���������
' - ToProcEnd(iProc)    - ���������������� �� <*>ProcEnd ������� ���������
' S ProcReset(Proc)     - ����� � ����� ������ �������� Proc
' - StepIn()            - ������ ���������� ����, �.�. ��������� ������ ������

Option Explicit

Const TRACE_STEP = "Trace"  ' ����������� ��� Trace ��� ������������� � �������
Public TraceStep As Boolean
Public TraceStop As Boolean
Public TraceWidth As Boolean

Sub ProcStart(Proc As String)
'
' - ProcStart(Proc) - ������ �������� Proc �� ������� Process � match.xlsm
'   7.8.12
'  26.8.12 - ������� ������������ ��������

    Dim Step As String, PrevStep As String
    Dim i As Integer
'---- �������������� ����� Trace
    TraceStep = False:    TraceStop = False:    TraceWidth = False
    
    Proc = Trim(Proc)
    
    i = ToStep(Proc)
    
    With DB_MATCH.Sheets(Process)
        .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 35
        Do While .Cells(i + 1, PROC_STEP_COL) <> PROC_END
            i = i + 1
            Step = .Cells(i, PROC_STEP_COL)
            If TraceStep Then
                .Activate
                .Rows(i).Select
            End If
            
            If .Cells(i, PROC_STEPDONE_COL) <> "1" Then
                PrevStep = .Cells(i, PROC_PREVSTEP_COL)
                If PrevStep <> "" Then _
                    If Not IsDone(Proc, PrevStep) Then GoTo Err
                    
                .Cells(1, PROCESS_NAME_COL) = Proc      '��� ��������
                .Cells(1, STEP_NAME_COL) = Step         '��� ����
                
'*************************************
                Exec Step, i        '*  ��������� ���
'*************************************
            
            End If
        Loop
        .Cells(1, PROCESS_NAME_COL) = "": .Cells(1, STEP_NAME_COL) = ""
        .Range(Cells(i + 1, 1), Cells(i + 1, 2)).Interior.ColorIndex = 35
    End With
''    MS "<*> ������� " & Proc & " ��������!"
    Exit Sub
Err:
    ErrMsg FATAL_ERR, "�������� ������������������ ����� ��������� " & Proc
    End
End Sub
Function IsDone(ByVal Proc As String, ByVal Step As String) As Boolean
'
' - IsDone(Proc, Step) - ��������, ��� ��� Step �������� ��� Proc ��������,
'                        � ���� �� �������� - ������ ����������� ��� ���������
'   7.8.12
'  16.8.12 - bug fix ��� PrevStep ��� ������� ����� ������

    Dim i As Integer
    Dim iStep As Long
    Dim S() As String   '=����� ���������� PrevStep, ����������� ","
    Dim X() As String   '=������ ����� ����� ���� ���� <Proc>/<Step>
    Dim Rep As String, Done As String
    
    Proc = Trim(Proc): Step = Trim(Step)
    
    If Step = REP_LOADED Then
        i = ToStep(Proc)
        Rep = DB_MATCH.Sheets(Process).Cells(i, PROC_REP1_COL)
        GetRep Rep
        If RepTOC.Made <> REP_LOADED Then
            Dim Msg As String
            ErrMsg FATAL_ERR, "IsDone: �� 'Loaded' ���� ��� �������� " _
                & Proc & " �� ���� " & Step & vbCrLf & vbCrLf _
                & "����� " & RepTOC.Name & " ���� ��������� ������!"
            Stop
            End
        Else
            If TraceStep Then MS "����� " & Rep & " ������������� 'Loaded'"
            IsDone = True
            Exit Function
        End If
    Else
        S = split(Trim(Step), ",")
        For i = LBound(S) To UBound(S)
            If InStr(S(i), "/") <> 0 Then
                X = split(S(i), "/")
                If Proc = X(0) Then ErrMsg FATAL_ERR, "����������� �������� � PrevStep!!"
                If Not IsDone(X(0), X(1)) Then ProcStart X(0)
            Else
                iStep = ToStep(Proc, S(i))
                If DB_MATCH.Sheets(Process).Cells(iStep, PROC_STEPDONE_COL) <> "" Then
                    IsDone = True
                    Exit Function
                End If
                ProcStart Proc  '����� - ������ �� PrevStep
            End If
        Next i
        IsDone = True
        Exit Function
    End If
End Function
Sub Exec(Step As String, iProc)
'
' - Exec(Step, iProc) - ����� ���� Step �� ������ iProc ������� ���������
'   7.8.12
'  26.8.12 - ������� ������ � Process ��� ������� ������������ ����
'   1.9.12 - ������� ����
       
    Dim Code As String
    Dim R As TOCmatch       '= �������������� �������� - �����
            
    If Step = PROC_END Or Step = "" Then Exit Sub
    
    With DB_MATCH.Sheets(Process)
'-- Trace - ����������� ��� ��� ������� ������������� � ������� �����
        If Not TraceStep Then TraceStep = False
        If Step = TRACE_STEP Then
            TraceStep = True
            TraceStop = False
            TraceWidth = False
            If .Cells(iProc, PROC_PAR1_COL) = 1 Then TraceStop = True
            If .Cells(iProc, PROC_PAR2_COL) = "W" Then TraceWidth = True
            Exit Sub
        End If

'*********** ����� ������������ - ���� ***********************
        Code = "'" & DirDBs & F_MATCH & "'!" & Step
        
        .Cells(1, STEP_NAME_COL) = Step
        If TraceStep Then
            MS "<> ������� " & .Cells(1, PROCESS_NAME_COL) _
                & " ����� ����������� ���� " & Step
            If TraceStop Then Stop
        End If
        ExRespond = True
        
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
        StepOut Step, iProc
    End With
End Sub
Sub StepIn()
'
' - StepIn()    - ������ ���������� ����, �.�. ��������� � ����� ������ ������
'   1.9.12

    Const FILE_PARAMS = 5   ' ������������ ���������� ������ � ����
    
    Dim iStep As Integer, i As Long
    Dim P As TOCmatch, S As TOCmatch, Rep As String
    
    ScreenUpdate False
    
    GetRep Process
    With DB_MATCH.Sheets(Process)
        PublicProcName = .Cells(1, PROCESS_NAME_COL)
        PublicStepName = .Cells(1, STEP_NAME_COL)
        
        iStep = ToStep(PublicProcName, PublicStepName)
        
        For i = FILE_PARAMS To 1 Step -1
            Rep = .Cells(iStep, i + PROC_REP1_COL - 1)
            If Rep <> "" Then
                S = GetRep(Rep)
                Workbooks(S.RepFile).Sheets(S.SheetN).Activate
            End If
        Next i
    End With
End Sub
Sub StepOut(Step As String, iProc)
'
' - StepOut()   - ���������� ���������� ���� � ������� � TOCmatch
'   8/10/12
'  28.10.12 - ����� ���������� ������ � TOCmatch �� ����������, �������������� � ����
'   9.11.12 - ��� ��������� ������?

    Dim Proc As String, R As TOCmatch
    
    ScreenUpdate True
    
    With DB_MATCH.Sheets(Process)
        Application.StatusBar = False
        .Activate
        .Cells(iProc, PROC_STEPDONE_COL) = "1"  ' Done = "1" - ��� ��������
        .Cells(iProc, PROC_TIME_COL) = Now
        .Range(Cells(iProc, 1), Cells(iProc, 3)).Interior.ColorIndex = 35
        .Cells(1, STEP_NAME_COL) = ""
        .Cells(1, 1) = Now
        If .Cells(iProc, PROC_REP1_COL) = "" Then Exit Sub  '��� ��������� ������?
        Proc = .Cells(1, PROCESS_NAME_COL)                  '��� �������� ������?
        If Proc = "" Then Exit Sub
        R = GetRep(.Cells(ToStep(Proc, Step), PROC_REP1_COL)) '�������������� ��������
        R.Made = Step
        R.Dat = Now
        RepTOC = R
        WrTOC
    End With
End Sub
Function ToStep(Proc, Optional Step As String = "") As Integer
'
' - ToStep(Proc, [Step]) - ���������� ����� ������ ������� ���������
'   7.8.12
    
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

MyProc: .Cells(1, PROCESS_NAME_COL) = Proc      '��� ��������
        .Cells(1, STEP_NAME_COL) = Step         '��� ����
        ToStep = i
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
Function ToProcEnd(ByVal iProc As Long) As Long
'
' - ToProcEnd(iProc)    - ���������������� �� <*>ProcEnd
' 30.9.12

    Dim P As TOCmatch
    
    P = GetRep(Process)
    ToProcEnd = iProc
    Do While DB_MATCH.Sheets(Process).Cells(ToProcEnd, PROC_STEP_COL) <> PROC_END
        ToProcEnd = ToProcEnd + 1
        If ToProcEnd >= P.EOL Then GoTo ErrExit
    Loop
    Exit Function
ErrExit:
    ErrMsg FATAL_ERR, "ToProcEnd> �� ������ ����� �������� �� ������ iProc=" & iProc
End Function
Sub WrProcResult(NewLine As Long)
'
' - WrProcResult(NewLine)   - ������ ���������� ���� � ������� PrevSter ��������
' 30.9.12

    Dim i As Long
    
    With DB_MATCH.Sheets(Process)
        i = ToStep(.Cells(1, PROCESS_NAME_COL))
        i = ToProcEnd(i)
    
        .Cells(i, PROC_PREVSTEP_COL) = NewLine
        .Cells(i, PROC_PREVSTEP_COL).Interior.Color = rgbGreen
    End With
End Sub
Sub ProcReset(Proc As String, _
    Optional ProcToReset As String = "", Optional StepToReset As String, Optional Col As Long)
'
' S ProcReset(Proc,[ProcToReset, StepToReset, Col]) - ����� � ����� ������ �������� Proc
' 1.10.12
' 11.11.12 - ������� ������ � ���� StepToReset � ������� Col

    Dim i As Long
    
    GetRep Process
    With DB_MATCH.Sheets(Process)
        If ProcToReset <> "" Then
            i = ToStep(ProcToReset, StepToReset)
            .Cells(i, Col) = ""
        End If
        i = ToStep(Proc)
        .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 0
        Do While .Cells(i, PROC_STEP_COL) <> PROC_END
            i = i + 1
            .Cells(i, PROC_STEPDONE_COL) = ""
            .Cells(i, PROC_TIME_COL) = ""
            .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 0
        Loop
    
        ProcStart Proc
    End With
End Sub
Sub CheckProc0(NewProcResult As String)
'
' S CheckProc0(NewProcResult)   - ��������, ��� ��������������� ������� �� �����
'                                 ����� "��������������" ������� � SF
' 1/10/12

    If NewProcResult <> "0" Then
        ErrMsg FATAL_ERR, PublicProcName & ": CheckProc0> � ���������� �� '0'"
        End
    End If
End Sub
