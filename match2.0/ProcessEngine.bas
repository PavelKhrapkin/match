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
' 14.9.12 �.�.�������
'
' - ProcStart(Proc)     - ������ �������� Proc �� ������� Process � match.xlsm
' - IsDone(Proc, Step)  - ��������, ��� ��� Step �������� Proc ��� ��������
' - Exec(Step, iProc)   - ����� ���� Step �� ������ iProc ������� ���������
' - ToStep(Proc,[Step]) - ���������� ����� ������ ������� ���������
' - StepIn()            - ������ ���������� ����, �.�. ��������� ������ ������
' S Adapt(F) - ��������� �������� �� ����� F
' - Adater(Request, X, F_rqst, IsErr) - ������������ X � �������� "Request"
'        � �������� ������� � ��������� F_rqst. IsErr=True - ������ � ��������
' - FetchDoc(F_rqst, X, IsErr) - ���������� ������ �� ���������� ���������
'                   �� ������� F_rqst ��� �������� ���� X. IsErr=True - ������

Option Explicit

Const TRACE_STEP = "Trace"  ' ����������� ��� Trace ��� ������������� � �������
Public TraceStep As Boolean
Public TraceStop As Boolean

Sub ProcStart(Proc As String)
'
' - ProcStart(Proc) - ������ �������� Proc �� ������� Process � match.xlsm
'   7.8.12
'  26.8.12 - ������� ������������ ��������

    Dim Step As String, PrevStep As String
    Dim i As Integer
    
    Proc = Trim(Proc)
    
    i = ToStep(Proc)
    
    With DB_MATCH.Sheets(Process)
        .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 35
        Do While Step <> PROC_END
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
        .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 35
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
Sub Exec(Step, iProc)
'
' - Exec(Step, iProc) - ����� ���� Step �� ������ iProc ������� ���������
'   7.8.12
'  26.8.12 - ������� ������ � Process ��� ������� ������������ ����
'   1.9.12 - ������� ����
       
    Dim Code As String, Proc As String
    Dim R As TOCmatch       '= �������������� �������� - �����
            
    If Step = PROC_END Or Step = "" Then Exit Sub
    
    With DB_MATCH.Sheets(Process)
'-- Trace - ����������� ��� ��� ������� ������������� � ������� �����
        If Not TraceStep Then TraceStep = False
        If Step = TRACE_STEP Then
            TraceStep = True
            TraceStop = False
            If .Cells(iProc, PROC_PAR1_COL) = 1 Then TraceStop = True
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
        Application.StatusBar = False
        .Activate
        .Cells(iProc, PROC_STEPDONE_COL) = "1"  ' Done = "1" - ��� ��������
        .Cells(iProc, PROC_TIME_COL) = Now
        .Range(Cells(iProc, 1), Cells(iProc, 3)).Interior.ColorIndex = 35
        .Cells(1, STEP_NAME_COL) = ""
        .Cells(1, 1) = Now
        Proc = .Cells(1, PROCESS_NAME_COL)              '��� ��������
        If Proc = "" Then Exit Sub
        R = GetRep(.Cells(ToStep(Proc), PROC_REP1_COL)) '�������������� ��������
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
Sub StepIn()
'
' - StepIn()    - ������ ���������� ����, �.�. ��������� � ����� ������ ������
'   1.9.12

    Const FILE_PARAMS = 5   '������������ ���������� ������ � ����
    
    Dim iStep As Integer, i As Long
    Dim P As TOCmatch, S As TOCmatch, Rep As String
    
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
Sub Adapt(F As String)
'
' S Adapt(F) - ��������� �������� �� ����� F � match.xlsm
'
' ����� F ����� ���:
'   �����   - ��������� �������. ����� ������������ � ������������� ����� InsMyCol
'   MyCol   - ������� � ��������� ����� �� ����� �� �����. ���� "V" - ������ �����
'   Width   - ������ �������
'   Columns - ����� ������� � �������� �����- ����� � ������ ���������� � ��������
'   ������� - ������- ����� ��������, ��������������� � = <�������� �� Columns>
'   Fetch   - ������ �������������� ���������� ��� �������� �� ������ ����������
'
' 12.9.12
' 14.9.12 - ���� ������� �� ����� �������� - ��������� �������� �� ���������

    StepIn
    
    Dim FF As Range     '= ����� F
    Dim R As TOCmatch
    Dim Rqst As String, F_rqst As String, iX As Long, IsErr As Boolean
    Dim X As String, Y As String
    Dim i As Long, Col As Long
    
    Set FF = DB_MATCH.Sheets(Header).Range(F)
    
    With ActiveSheet
        R = GetRep(.Name)
        For i = 2 To R.EOL
            Progress i / R.EOL
            For Col = 1 To FF.Columns.Count
                iX = FF(4, Col)
                If iX > 0 Then
                    X = .Cells(i, iX)
                    Rqst = FF.Cells(5, Col)
                    F_rqst = FF.Cells(6, Col)
                    
                    Y = Adapter(Rqst, X, F_rqst, IsErr)
                    
                    If Not IsErr Then .Cells(i, Col) = Y
                End If
            Next Col
        Next i
    End With
End Sub
Function Adapter(Request, ByVal X, F_rqst, IsErr) As String
'
' - Adater(Request, X, F_rqst, IsErr) - ������������ X � �������� "Request"
'    � �������� ������� � ��������� F_rqst. IsErr=True - ������ � ��������
' 4.9.12
' 6.9.12 - bug fix

    Dim FF() As String, Tmp() As String, Cols() As String
    Dim Doc As String, C1 As Long, C2 As Long, Rng As Range
    Dim F() As String
    Dim i As Long, Par() As String
    
    IsErr = False
    X = Trim(X)
    
'--- ������ ������ �������� ���� <���>/C1,C2,C3...
    Dim AdapterName As String
    AdapterName = ""
    If Request <> "" Then
        Tmp = split(Request, "/")
        AdapterName = Tmp(0)
        If InStr(Request, "/") <> 0 Then Par = split(Tmp(1), ",")
    End If

'========== ������������� �������� =========
    Select Case AdapterName
    Case "MainContract":
        X = Trim(Replace(X, "�������", ""))
    Case "<>0":
        If X = "0" Then X = ""
    End Select
    
'--- FETCH ������ ������ ���������� �� ���������� ���� <Doc1>/C1:C2,<Doc2>/C1:C2,...
    If F_rqst <> "" And X <> "" Then
        
        FF = split(F_rqst, ",")
        For i = LBound(FF) To UBound(FF)
            X = FetchDoc(FF(i), X, IsErr)
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' ������ ������������ ������ ���� ��������� �� ����������� �� Doc ��������.
' � ���������� ���� ������������ ������ x(1 to 5) � ���������� � Fetch ��������� ���
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Next i
    End If

'******* ���������� �������� � ����������� Par ******
    Select Case AdapterName
    Case "", "MainContract": Adapter = X
    Case "��", "��������_�_SF", "Vendor":
        On Error GoTo AdapterFailure
        Adapter = WorksheetFunction.VLookup(X, DB_MATCH.Sheets("We").Range(AdapterName), Par(0), False)
        On Error GoTo 0
    Case "Dec": Adapter = Dec(X)
    Case "CurISO":
        Adapter = CurISO(X)
    Case "CurRate": Adapter = CurRate(X)
    Case "����":
        If X = "" Then
            Adapter = ""
        Else
            Adapter = DDMMYYYY(X)
        End If
    Case Else
        ErrMsg FATAL_ERR, "Adapter> �� ���������� " & AdapterName
    End Select
    Exit Function
    
AdapterFailure:
    ErrMsg WARNING, "������� " & AdapterName & "(" & X & ") �� ������� ������"
    IsErr = True
End Function
Function FetchDoc(F_rqst, X, IsErr) As String
'
' - FetchDoc(F_rqst, X, IsErr) - ���������� ������ �� ���������� ���������
'                   �� ������� F_rqst ��� �������� ���� X. IsErr=True - ������
'
' * F_rqst ����� ��� <Doc>/C1[:C2][/W]
' * <Doc>   - ��� ���������, ����� ����������� ������
' *   /     - �������� ������ ����������. �������� ��������� �����.
' *   :     - ��������� ��������� ������ ������
' *             ������ ������ - ���������� ��������� ��� ���������� �1[:�2]
' *  C1                 ���� ���� ������ �1 - ����������� ���� ����� �1
' * C1:C2               ���� �1:�2 - Lookup �� �1 -> �� C2 � Range �� Doc
' *             ������ ������ - ��������� ��������� ������ Fetch - /W ��� /0
' *  /W             - WARNING � Log, ��������� IsErr=False, ���� ��������� ""
' *  /0             - "" ������ ��������� (��������, ������� � ������)
' *  /D             - Default - "" ��������, �� IsErr=True ��� �������� �� ���������
' *             ������ ������ ����������� - �������� Log � IsErr = True
'
' 5.9.12
' 14.9.12 - �������� /D ��� ������ ������ - "�� ���������"

    FetchDoc = ""
    If F_rqst = "" Or X = "" Then GoTo ErrExit
        
    Dim Tmp() As String, Cols() As String, S As String
    Dim Doc As String, C1 As Long, C2 As Long, Rng As Range
            
    Tmp = split(F_rqst, "/")
    Doc = Tmp(0)
    Cols = split(Tmp(1), ":")
    C1 = Cols(0)
    
    Dim Rdoc As TOCmatch, W As Workbook
    Rdoc = GetRep(Doc)
    
    If UBound(Cols) < 1 Then
'--- �������� �1 - � ������ ���� �������� - ��������� �������� �� �������
        Dim Indx As Long
        Indx = X
'!!!!!!!!!!!!!!!!!!!!!!!!!!!
' ������ Indx=� - ��� ������ �����, �� � ���������� ��� ���� Split
'!!!!!!!!!!!!!!!!!!!!!!!!!!!
        If Indx <= 0 Then
            ErrMsg WARNING, "FetchDoc: " & Doc & "(" & Indx & "," & C1 _
                & ") - ������������ ����� ������"
            GoTo ErrExit
        End If
        S = Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN).Cells(Indx, C1)
    Else
'--- �������� �1:C2 - � ������ 2 ��������� - ��������� �������� �� Lookup
        C2 = Cols(1)
        Dim Lit As String
        Const A = 64            ' String("A")-1
        Lit = Chr(C1 + A) & ":" & Chr(C2 + A)
        Set Rng = Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN).Range(Lit)
        
        S = ""
        On Error Resume Next
        S = WorksheetFunction.VLookup(X, Rng, C2 - C1 + 1, False)
        On Error GoTo 0
    End If
'--- ��������� ������ 2 -- ���� S=""
    If S = "" Then
        If UBound(Tmp) >= 2 Then
            If Tmp(2) = "W" Then
                ErrMsg WARNING, "�������> ������ " & F_rqst _
                    & "(" & X & ") �� ��������, ��������� <�����>"
            End If
            If Tmp(2) <> "0" Then GoTo ErrExit
        Else
            ErrMsg WARNING, "�������> ������ " & F_rqst _
               & "(" & X & ") �� ��������, ��������� <�����>"
            GoTo ErrExit
        End If
    Else
        FetchDoc = S
    End If
    
OK_Exit:    IsErr = False
    Exit Function
ErrExit:    IsErr = True

End Function
