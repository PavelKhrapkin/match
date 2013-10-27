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
' 27.10.13 �.�.�������, �.����
'
' S/- ProcStart(Proc)   - ������ �������� Proc �� ������� Process � match.xlsm
' - IsDone(Proc, Step)  - ��������, ��� ��� Step �������� Proc ��� ��������
' - Exec(Step, iProc)   - ����� ���� Step �� ������ iProc ������� ���������
' - ToStep(Proc,[Step]) - ���������� ����� ������ ������� ���������
' - ToProcEnd(iProc)    - ���������������� �� <*>ProcEnd ������� ���������
' S ProcReset(Proc)     - ����� � ����� ������ �������� Proc
' - StepIn()            - ������ ���������� ����, �.�. ��������� ������ ������
' - StepOut()           - ���������� ���������� ���� � ������� � TOCmatch
' S MergeReps()         - ������� "������" ������� � ��������� "_OLD" � "Update"
' -DocReset(DocName)    - ����� ���� �����, ���������� � DocName
' - StepReset(iStep)    - ����� ���� � ������� ��������� - ��������!

Option Explicit

'========== ��������� ��������� ==================
Const TRACE_STEP = "Trace"  ' ����������� ��� Trace ��� ������������� � �������
Public TraceStep As Boolean
Public TraceStop As Boolean
Public TraceWidth As Boolean
'========== ���� ��������� ==================
Dim ProcStack As Collection

Sub ProcStart(Proc As String)
'
' - ProcStart(Proc) - ������ �������� Proc �� ������� Process � match.xlsm
'   7.8.12
'  26.8.12 - ������� ������������ ��������
'  24.8.13 - �� ���������� �������� ���������� <*>ProcEnd � ��� ���������
'  30.8.13 - ����� �� PROC_END ��� ���������
'  26.10.13 - �� <*>ProcEnd ���������� ��� ����, ������� ���������� �������� ���������
'  27.10.13 - ������������ ��������� ���������� ���������, �������������� � ���� �����

    Dim Step As String, PrevStep As String
    Dim i As Integer, Doc As String, � As TOCmatch
'---- �������������� ����� Trace
    TraceStep = False:    TraceStop = False:    TraceWidth = False
    
    Proc = Trim(Proc)
    
    On Error GoTo ProcStackInit
ProcAdd:  ProcStack.Add Proc
    On Error GoTo 0
    
    i = ToStep(Proc)
    
    With DB_MATCH.Sheets(Process)
        .Activate
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
        
        Dim ProcEndLine As Long: ProcEndLine = i + 1
        For i = 0 To 5
            DocReset .Cells(ProcEndLine, i + PROC_REP1_COL)
        Next i
        ProcStack.Remove ProcStack.Count
        .Activate
        .Cells(1, PROCESS_NAME_COL) = "": .Cells(1, STEP_NAME_COL) = ""
        .Range(Cells(ProcEndLine, 1), Cells(ProcEndLine, 2)).Interior.ColorIndex = 35
        i = ToStep(Proc)
        Doc = .Cells(i, PROC_REP1_COL)
        If Doc = "" Then GoTo Ex    '���� ������� �� ����������� ������� �������� -> �����
        RepTOC = GetRep(Doc)
        RepTOC.Made = PROC_END
        WrTOC
''        MS "<*> ������� " & Proc & " ��������!"
    End With
Ex: Exit Sub
Err:
    ErrMsg FATAL_ERR, "�������� ������������������ ����� ��������� " & Proc
    End
ProcStackInit:  Set ProcStack = New Collection
    GoTo ProcAdd
End Sub
Function IsDone(ByVal Proc As String, ByVal Step As String) As Boolean
'
' - IsDone(Proc, Step) - ��������, ��� ��� Step �������� ��� Proc ��������,
'                        � ���� �� �������� - ������ ����������� ��� ���������
'   7.8.12
'  16.8.12 - bug fix ��� PrevStep ��� ������� ����� ������
'  27.8.13 - �� ���������� ���������� ��������� RepTOC

    Dim i As Integer
    Dim iStep As Long
    Dim S() As String   '=����� ���������� PrevStep, ����������� ","
    Dim X() As String   '=������ ����� ����� ���� ���� <Proc>/<Step>
    Dim Rep As String, Done As String
    Dim Report As TOCmatch
    
    Proc = Trim(Proc): Step = Trim(Step)
    
    If Step = REP_LOADED Then
        i = ToStep(Proc)
        Rep = DB_MATCH.Sheets(Process).Cells(i, PROC_REP1_COL)
        Report = GetRep(Rep)
        If Report.Made <> REP_LOADED Then
            Dim msg As String
            ErrMsg FATAL_ERR, "IsDone: �� 'Loaded' ���� ��� �������� " _
                & Proc & " �� ���� " & Step & vbCrLf & vbCrLf _
                & "����� " & Report.Name & " ���� ��������� ������!"
            Stop
            End
        Else
            If TraceStep Then MS "����� " & Rep & " ������������� 'Loaded'"
            IsDone = True
            Exit Function
        End If
    Else
        S = Split(Trim(Step), ",")
        For i = LBound(S) To UBound(S)
            If InStr(S(i), "/") <> 0 Then
                X = Split(S(i), "/")
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
Function IsStepDone(ByVal Proc As String, ByVal Step As String) As Boolean
'
' - IsStepDone(Proc, Step) - ��������, ��� ��� Step �������� ��� Proc ��������
' 23.10.12

    Dim i As Long
    
    IsStepDone = True
    If Step = REP_LOADED Then Exit Function
    i = ToStep(Trim(Proc), Trim(Step))
    If DB_MATCH.Sheets(Process).Cells(i, PROC_STEPDONE_COL) <> "1" Then IsStepDone = False
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
'  11.08.13 - ���������� EOL ��������������� ��������� � TOCmatch
'  26.08.13 - ���� ��� ����� RepTOC.EOL, ����� ���������� TOC �� �������� ������
'  15.09.13 - ��������� ������������ ��� ProcReset ������ ����

    Dim Proc As String, Doc As String, i As Long
    
    ScreenUpdate True
'    If Step = "ProcStart" Then Exit Sub
    RepTOC.Made = Step: RepTOC.Dat = Now
    Doc = DB_MATCH.Sheets(Process).Cells(iProc, PROC_REP1_COL)
    WrTOC Doc     ' ��������� EOL � TOC � ��������, ��� �� ������� ������
    
    With DB_MATCH.Sheets(Process)
        Application.StatusBar = False
        .Activate
        If Step <> "ProcReset" _
                Or .Cells(iProc, PROC_PAR1_COL) <> .Cells(1, PROCESS_NAME_COL) Then
            .Cells(iProc, PROC_STEPDONE_COL) = "1"  ' Done = "1" - ��� ��������
        End If
        .Cells(iProc, PROC_TIME_COL) = Now
        .Range(Cells(iProc, 1), Cells(iProc, 3)).Interior.ColorIndex = 35
        .Cells(1, STEP_NAME_COL) = ""
        .Cells(1, 1) = Now
''''''        Proc = .Cells(1, PROCESS_NAME_COL)                  '��� �������� ������?
''''''        If Proc = "" Then Exit Sub
''''''        Doc = .Cells(ToStep(Proc, Step), PROC_REP1_COL)     '�������������� ��������
''''''        If Doc = "" Then Exit Sub                           '��� ��������� ������?
    End With
End Sub
Function ToStep(Proc, Optional Step As String = "") As Integer
'
' - ToStep(Proc, [Step]) - ���������� ����� ������ ������� ���������
'   7.8.12
'  11.8.13 - ����� ��������� ��������� �� ������
'  27.8.13 - �� ���������� ���������� ��������� RepTOC
    
    Dim P As TOCmatch           '������ ������� ��������� � ���� TOCmatch
    Dim StepName As String      '=��� �������� ����
    Dim ProcName As String      '=��� �������� ��������
    Dim i As Integer
    
    P = GetRep(Process)
    
    With DB_MATCH.Sheets(Process)
        For i = 6 To P.EOL
            ProcName = .Cells(i, PROC_NAME_COL)
            StepName = .Cells(i, PROC_STEP_COL)
            If StepName = PROC_START And ProcName = Proc Then GoTo MyProc
        Next i
        ErrMsg FATAL_ERR, "ToProc: �� ������ ������� " & Proc _
                        & vbCrLf & "������� ��������� ������� Process � �� EOL � TOCmatch."
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
' 15.09.13 - ��������� ������������ ��� ProcReset ������ ����

    Dim i As Long, IsMe As Boolean
    IsMe = False
    
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
            If .Cells(i, PROC_STEP_COL) = "ProcReset" _
                    And .Cells(i, PROC_PAR1_COL) = Proc Then IsMe = True
        Loop
        If Not IsMe Then ProcStart Proc
    End With
End Sub
Sub MergeReps()
'
' S MergeReps()    - ������� "������" ������� � ��������� "_OLD" � "Update"
'
'   * �������� ��� �������� � ��������� 1�
'
' 24.8.13
'  7.9.13 - bug fix - ������ �������� ��� ������� ������ �� �����
'  6.10.13 - ���������� ������ �� �������� ��������� � �����

    Dim RefSummary As String
    Dim R As TOCmatch
    Dim OldRepName As String
    Dim RoldEOL As Long, Col As Long, i As Long, FrRow As Long, ToRow As Long
    Dim FrDate As Date, ToDate As Date
    Dim FrDateRow, ToDateRow
    
    StepIn
    
    RepName = ActiveSheet.Name
    R = GetRep(RepName)
    OldRepName = RepName & "_OLD"
    If Not SheetExists(OldRepName) Then Exit Sub
    R.EOL = EOL(RepName) - GetReslines(RepName)
    RoldEOL = EOL(OldRepName) - GetReslines(RepName)
    
'-- ���� ��������� - ������ TOC �� ������ ������
    With DB_MATCH.Sheets(TOC)
        FrDateRow = .Cells(R.iTOC, TOC_FRDATEROW_COL)
        ToDateRow = .Cells(R.iTOC, TOC_TODATEROW_COL)
        Col = R.MyCol + .Cells(R.iTOC, TOC_DATECOL_COL)
        FrDate = .Cells(R.iTOC, TOC_NEW_FRDATE_COL)
        ToDate = .Cells(R.iTOC, TOC_NEW_TODATE_COL)
        RefSummary = .Cells(R.iTOC, TOC_FORMSUMMARY)
    End With
    
    With Workbooks(R.RepFile).Sheets(OldRepName)
        .Activate
'-- ��������� ��������_OLD
        For i = 1 To BIG
            If ActiveWorkbook.Sheets(i).Name = OldRepName Then Exit For
        Next i
        SheetSort i, Col
        
        FrRow = 0: ToRow = 0
        For i = 2 To RoldEOL
            If .Cells(i, Col) >= FrDate And FrRow = 0 Then FrRow = i
            If .Cells(i, Col) <= ToDate Then
                ToRow = i
            Else
                GoTo InsRow
            End If
        Next i
        ToRow = RoldEOL + 1
'-- �������� ������� �������� (_OLD) �� ������ FrRow � �� ToRow �� EOL � ����� ��������
InsRow: If FrRow = 0 Then FrRow = RoldEOL + 1
        If FrRow > 2 Then
            .Rows("2:" & FrRow - 1).Copy
            Workbooks(R.RepFile).Sheets(R.SheetN).Rows("2:2").Insert Shift:=xlDown
        End If
        ToRow = ToRow + 1
        If ToRow < RoldEOL Then
            .Rows(ToRow & ":" & RoldEOL).Copy
            Workbooks(R.RepFile).Sheets(R.SheetN).Rows(R.EOL & ":" & R.EOL).Insert Shift:=xlDown
        End If
    End With
        
'-- ���������� ����� � ����� � ���������� ������� ��������
    With Workbooks(R.RepFile)
        With .Sheets(R.SheetN)
            R.EOL = EOL(RepName, Workbooks(R.RepFile)) - GetReslines(RepName)
            DB_MATCH.Sheets(Header).Range(RefSummary).Copy _
                Destination:=.Cells(R.EOL + 2, 1)
            SheetSort i, Col
            If ToDateRow = "EOL" Then ToDateRow = R.EOL
            FrDate = .Cells(FrDateRow, Col) ' ������������ FrDate � ToDate
            ToDate = .Cells(ToDateRow, Col)
        End With
        Application.DisplayAlerts = False
            .Sheets(OldRepName).Delete
        Application.DisplayAlerts = True
    End With
    
'---- ������������ FrDate � ToDate � TOCmatch
    With DB_MATCH.Sheets(TOC)
        .Cells(R.iTOC, TOC_FRDATE_COL) = FrDate
        .Cells(R.iTOC, TOC_TODATE_COL) = ToDate
    End With
End Sub
Sub DocReset(DocName As String)
'
' -DocReset(DocName)    - ����� ���� �����, ���������� � DocName
'
' 26.10.13
' 27.10.13 - �������� �� ������ ��������� � � ����������� ������ ���������

    Dim i As Long, Proc As String, P
    If DocName = "" Then GoTo Ex
    
    Dim LocalTOC As TOCmatch
    LocalTOC = GetRep(DocName)
    
    If SheetExists(DocName & "_OLD") Then GoTo Ex
    If LocalTOC.ChkSum = DocCheckSum(DocName) Then GoTo Ex
    
    Dim ChkStack As Boolean: ChkStack = True
    If ProcStack Is Nothing Then ChkStack = False
    
    With DB_MATCH.Sheets(Process)
        For i = 6 To EOL(Process, DB_MATCH)
            If .Cells(i, PROC_COMMENT_COL) <> "" Then GoTo NxtI
            If .Cells(i, PROC_STEP_COL) = PROC_START Then Proc = .Cells(i, PROC_NAME_COL)
            If .Cells(i, PROC_REP1_COL) = DocName _
                    Or .Cells(i, PROC_REP1_COL + 1) = DocName _
                    Or .Cells(i, PROC_REP1_COL + 2) = DocName _
                    Or .Cells(i, PROC_REP1_COL + 3) = DocName _
                    Or .Cells(i, PROC_REP1_COL + 4) = DocName Then
                    
                If ChkStack Then
                    For Each P In ProcStack
                        If P = Proc Then GoTo NxtI
                    Next P
                End If
                StepReset i
            End If
NxtI:   Next i
    End With
Ex: Exit Sub
End Sub
Sub StepReset(iStep)
'
' - StepReset(iStep) - ����� ���� � ������� ��������� - ��������!
' 28.8.12
'  9.9.12 - bug fix � ����� ������������ ���� ��� �������� ������ ���������
' 13.9.12 - bug fix - �� ���������� ���� <*>ProcStart
' 22.10.13 - bug fix - Range ������� 1..3 ���������
' 23.10.13 - ���������� �� ������ ��� iStep, � ���� ������� ��������� � ������
'            � ��� ����, ��� ������� iStep �������� "����������" - PrevStep
' 24.10.13 - bug fix - ����� ��������� ������ <*>ProcEnd PrevStep

    Dim i As Integer, iProc As Integer
    Dim ThisProc As String, PrevSteps() As String
    Dim PrevS, PrS() As String
    
    With DB_MATCH.Sheets(Process)
        If .Cells(iStep, PROC_STEPDONE_COL) = "" Then Exit Sub
'---- ����� ����� �� iStep �� <*>ProcEnd � ������� ������ ��������� "<*>ProcStart"
        For i = 6 To EOL(Process, DB_MATCH)
            If .Cells(i, PROC_STEP_COL) = PROC_START Then iProc = i
            If i >= iStep Then
                .Cells(i, PROC_STEPDONE_COL) = ""
                .Range("A" & i & ":C" & i).Interior.ColorIndex = 0
            End If
            If .Cells(i, PROC_STEP_COL) = PROC_END And i > iStep Then Exit For
        Next i
        .Range("A" & iProc & ":C" & iProc).Interior.ColorIndex = 0
'---- ����� ���� �����, � ������� � PrevStep ��������� �� ������������� ����
        For i = 6 To EOL(Process, DB_MATCH)
            If .Cells(i, PROC_STEP_COL) = PROC_START Then ThisProc = .Cells(i, PROC_NAME_COL)
            If .Cells(i, PROC_STEP_COL) <> PROC_END _
                    And .Cells(i, PROC_COMMENT_COL) <> "" Then
                PrevSteps = Split(.Cells(i, PROC_PREVSTEP_COL), ",")
                For Each PrevS In PrevSteps
                    If InStr(PrevS, "/") = 0 Then
                       If Not IsStepDone(ThisProc, PrevS) Then StepReset i  ' ��������!
                    Else
                       PrS = Split(PrevS, "/")
                       If Not IsStepDone(PrS(0), PrS(1)) Then StepReset i   ' ��������!
                    End If
                Next PrevS
            End If
        Next i
    End With
End Sub
