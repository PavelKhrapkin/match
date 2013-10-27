Attribute VB_Name = "ProcessEngine"
'---------------------------------------------------------------------------------------
' ProcessEngine - процессор Процессов.
'         * Состояние и Шаги процессов записаны в Process файла match.xlsm
'         * Процессы Loader'ы запускаются при загрузке в DBs нового Документа.
'         * Загрузка нового Документа сбрасывает в начальное состояние все
'           Процессы,связанные с использованием данного Документа
'         * Процессы Handler'ы и Reporter'ы инициируются из меню формы Главной Панели
'         * Все Процессы повторновходимы. Если Шаг уже выполнен - он пропускается.
'         * Перед выполнением Шага проверяется поле Done по шагу PrevStep.
'           PrevStep может иметь вид <другой Процесс> / <Шаг>.
'
' 27.10.13 П.Л.Храпкин, А.Пасс
'
' S/- ProcStart(Proc)   - запуск Процесса Proc по таблице Process в match.xlsm
' - IsDone(Proc, Step)  - проверка, что шаг Step процесса Proc уже выполнен
' - Exec(Step, iProc)   - вызов Шага Step по строке iProc таблицы Процессов
' - ToStep(Proc,[Step]) - возвращает номер строки таблицы Процессов
' - ToProcEnd(iProc)    - позиционирование на <*>ProcEnd таблицы Процессов
' S ProcReset(Proc)     - сброс и новый запуск Процесса Proc
' - StepIn()            - начало исполнения Шага, т.е. активация нужных листов
' - StepOut()           - завершение выполнения Шага с записью в TOCmatch
' S MergeReps()         - слияние "полных" отчетов в суффиксом "_OLD" и "Update"
' -DocReset(DocName)    - сброс всех шагов, работающих с DocName
' - StepReset(iStep)    - сброс Шага в таблице Процессов - РЕКУРСИЯ!

Option Explicit

'========== Константы Процессов ==================
Const TRACE_STEP = "Trace"  ' специальный Шаг Trace для трассирования и отладки
Public TraceStep As Boolean
Public TraceStop As Boolean
Public TraceWidth As Boolean
'========== Стек Процессов ==================
Dim ProcStack As Collection

Sub ProcStart(Proc As String)
'
' - ProcStart(Proc) - запуск Процесса Proc по таблице Process в match.xlsm
'   7.8.12
'  26.8.12 - окраска выполненного Процесса
'  24.8.13 - по завершению Процесса записываем <*>ProcEnd в ТОС Документа
'  30.8.13 - выход по PROC_END без Документа
'  26.10.13 - по <*>ProcEnd сбрасываем все Шаги, которые используют выходные Документы
'  27.10.13 - понадобилась коллекция запущенных Процессов, организованных в виде стека

    Dim Step As String, PrevStep As String
    Dim i As Integer, Doc As String, К As TOCmatch
'---- инициализируем флаги Trace
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
                    
                .Cells(1, PROCESS_NAME_COL) = Proc      'имя Процесса
                .Cells(1, STEP_NAME_COL) = Step         'имя Шага
                
'*************************************
                Exec Step, i        '*  выполняем Шаг
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
        If Doc = "" Then GoTo Ex    'если Процесс не обрабатывал никакой Документ -> выход
        RepTOC = GetRep(Doc)
        RepTOC.Made = PROC_END
        WrTOC
''        MS "<*> Процесс " & Proc & " завершен!"
    End With
Ex: Exit Sub
Err:
    ErrMsg FATAL_ERR, "Нарушена последовательность шагов процедуры " & Proc
    End
ProcStackInit:  Set ProcStack = New Collection
    GoTo ProcAdd
End Sub
Function IsDone(ByVal Proc As String, ByVal Step As String) As Boolean
'
' - IsDone(Proc, Step) - проверка, что шаг Step процесса уже Proc выполнен,
'                        а если не выполнен - запуск исполняющей его Процедуры
'   7.8.12
'  16.8.12 - bug fix про PrevStep без запятой давал ошибку
'  27.8.13 - не используем глобальную структуру RepTOC

    Dim i As Integer
    Dim iStep As Long
    Dim S() As String   '=части требований PrevStep, разделенные ","
    Dim X() As String   '=каждая часть может быть вида <Proc>/<Step>
    Dim Rep As String, Done As String
    Dim Report As TOCmatch
    
    Proc = Trim(Proc): Step = Trim(Step)
    
    If Step = REP_LOADED Then
        i = ToStep(Proc)
        Rep = DB_MATCH.Sheets(Process).Cells(i, PROC_REP1_COL)
        Report = GetRep(Rep)
        If Report.Made <> REP_LOADED Then
            Dim msg As String
            ErrMsg FATAL_ERR, "IsDone: Не 'Loaded' файл для Процесса " _
                & Proc & " на Шаге " & Step & vbCrLf & vbCrLf _
                & "Отчет " & Report.Name & " надо загрузить заново!"
            Stop
            End
        Else
            If TraceStep Then MS "Отчет " & Rep & " действительно 'Loaded'"
            IsDone = True
            Exit Function
        End If
    Else
        S = Split(Trim(Step), ",")
        For i = LBound(S) To UBound(S)
            If InStr(S(i), "/") <> 0 Then
                X = Split(S(i), "/")
                If Proc = X(0) Then ErrMsg FATAL_ERR, "Бесконечная рекурсия в PrevStep!!"
                If Not IsDone(X(0), X(1)) Then ProcStart X(0)
            Else
                iStep = ToStep(Proc, S(i))
                If DB_MATCH.Sheets(Process).Cells(iStep, PROC_STEPDONE_COL) <> "" Then
                    IsDone = True
                    Exit Function
                End If
                ProcStart Proc  'здесь - только по PrevStep
            End If
        Next i
        IsDone = True
        Exit Function
    End If
End Function
Function IsStepDone(ByVal Proc As String, ByVal Step As String) As Boolean
'
' - IsStepDone(Proc, Step) - проверка, что шаг Step процесса уже Proc выполнен
' 23.10.12

    Dim i As Long
    
    IsStepDone = True
    If Step = REP_LOADED Then Exit Function
    i = ToStep(Trim(Proc), Trim(Step))
    If DB_MATCH.Sheets(Process).Cells(i, PROC_STEPDONE_COL) <> "1" Then IsStepDone = False
End Function

Sub Exec(Step As String, iProc)
'
' - Exec(Step, iProc) - вызов Шага Step по строке iProc таблицы Процессов
'   7.8.12
'  26.8.12 - окраска строки в Process для успешно выполненного Шага
'   1.9.12 - ревизия кода
       
    Dim Code As String
    Dim R As TOCmatch       '= обрабатываемый Документ - отчет
            
    If Step = PROC_END Or Step = "" Then Exit Sub
    
    With DB_MATCH.Sheets(Process)
'-- Trace - специальный Шаг для запуска трассирования и отладки Шагов
        If Not TraceStep Then TraceStep = False
        If Step = TRACE_STEP Then
            TraceStep = True
            TraceStop = False
            TraceWidth = False
            If .Cells(iProc, PROC_PAR1_COL) = 1 Then TraceStop = True
            If .Cells(iProc, PROC_PAR2_COL) = "W" Then TraceWidth = True
            Exit Sub
        End If

'*********** вызов подпрограммы - Шага ***********************
        Code = "'" & DirDBs & F_MATCH & "'!" & Step
        
        .Cells(1, STEP_NAME_COL) = Step
        If TraceStep Then
            MS "<> Процесс " & .Cells(1, PROCESS_NAME_COL) _
                & " перед выполнением Шага " & Step
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
'-- запись отметки о Шаге в TOCmatch и в таблицу Процессов
        StepOut Step, iProc
    End With
End Sub
Sub StepIn()
'
' - StepIn()    - начало исполнения Шага, т.е. активация и выбор нужных листов
'   1.9.12

    Const FILE_PARAMS = 5   ' максимальное количество файлов в Шаге
    
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
' - StepOut()   - завершение выполнения Шага с записью в TOCmatch
'   8/10/12
'  28.10.12 - более аккуратная работа с TOCmatch по документам, обрабатываемым в Шаге
'   9.11.12 - имя Документа пустое?
'  11.08.13 - Записываем EOL обрабатываемого документа в TOCmatch
'  26.08.13 - Если Шаг менял RepTOC.EOL, нужно переписать TOC до проверки Штампа
'  15.09.13 - исключаем зацикливание при ProcReset самого себя

    Dim Proc As String, Doc As String, i As Long
    
    ScreenUpdate True
'    If Step = "ProcStart" Then Exit Sub
    RepTOC.Made = Step: RepTOC.Dat = Now
    Doc = DB_MATCH.Sheets(Process).Cells(iProc, PROC_REP1_COL)
    WrTOC Doc     ' перепишем EOL в TOC и проверим, что не сбились Штампы
    
    With DB_MATCH.Sheets(Process)
        Application.StatusBar = False
        .Activate
        If Step <> "ProcReset" _
                Or .Cells(iProc, PROC_PAR1_COL) <> .Cells(1, PROCESS_NAME_COL) Then
            .Cells(iProc, PROC_STEPDONE_COL) = "1"  ' Done = "1" - Шаг выполнен
        End If
        .Cells(iProc, PROC_TIME_COL) = Now
        .Range(Cells(iProc, 1), Cells(iProc, 3)).Interior.ColorIndex = 35
        .Cells(1, STEP_NAME_COL) = ""
        .Cells(1, 1) = Now
''''''        Proc = .Cells(1, PROCESS_NAME_COL)                  'имя Процесса пустое?
''''''        If Proc = "" Then Exit Sub
''''''        Doc = .Cells(ToStep(Proc, Step), PROC_REP1_COL)     'обрабатываемый Документ
''''''        If Doc = "" Then Exit Sub                           'имя Документа пустое?
    End With
End Sub
Function ToStep(Proc, Optional Step As String = "") As Integer
'
' - ToStep(Proc, [Step]) - возвращает номер строки таблицы Процессов
'   7.8.12
'  11.8.13 - более подробное сообщение об ошибке
'  27.8.13 - не используем глобальную структуру RepTOC
    
    Dim P As TOCmatch           'строка таблицы Процессов в виде TOCmatch
    Dim StepName As String      '=Имя текущего Шага
    Dim ProcName As String      '=Имя текущего Процесса
    Dim i As Integer
    
    P = GetRep(Process)
    
    With DB_MATCH.Sheets(Process)
        For i = 6 To P.EOL
            ProcName = .Cells(i, PROC_NAME_COL)
            StepName = .Cells(i, PROC_STEP_COL)
            If StepName = PROC_START And ProcName = Proc Then GoTo MyProc
        Next i
        ErrMsg FATAL_ERR, "ToProc: Не найден Процесс " & Proc _
                        & vbCrLf & "Следует проверить таблицу Process и ее EOL в TOCmatch."
        Stop
        End

MyProc: .Cells(1, PROCESS_NAME_COL) = Proc      'имя Процесса
        .Cells(1, STEP_NAME_COL) = Step         'имя Шага
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
    ErrMsg FATAL_ERR, "ToStep: Обращение к несуществующему Шагу " & Step _
        & " Процесса " & Proc
End Function
Function ToProcEnd(ByVal iProc As Long) As Long
'
' - ToProcEnd(iProc)    - позиционирование на <*>ProcEnd
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
    ErrMsg FATAL_ERR, "ToProcEnd> не достиг конца Процесса со строки iProc=" & iProc
End Function
Sub WrProcResult(NewLine As Long)
'
' - WrProcResult(NewLine)   - запись результата Шага в колонку PrevSter Процесса
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
' S ProcReset(Proc,[ProcToReset, StepToReset, Col]) - сброс и новый запуск Процесса Proc
' 1.10.12
' 11.11.12 - очистка ячейки в Шаге StepToReset в колонке Col
' 15.09.13 - исключаем зацикливание при ProcReset самого себя

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
' S MergeReps()    - слияние "полных" отчетов в суффиксом "_OLD" и "Update"
'
'   * Отлажено для Платежей и Договоров 1С
'
' 24.8.13
'  7.9.13 - bug fix - всегда заменяем низ старого отчета до конца
'  6.10.13 - записываем строки из прежнего документа в новый

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
    
'-- куда вставлять - чтение TOC по НОВОМУ отчету
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
'-- сортируем документ_OLD
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
'-- копируем прежний документ (_OLD) от строки FrRow и от ToRow до EOL в новый документ
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
        
'-- дописываем пятку в новый и уничтожаем прежний документ
    With Workbooks(R.RepFile)
        With .Sheets(R.SheetN)
            R.EOL = EOL(RepName, Workbooks(R.RepFile)) - GetReslines(RepName)
            DB_MATCH.Sheets(Header).Range(RefSummary).Copy _
                Destination:=.Cells(R.EOL + 2, 1)
            SheetSort i, Col
            If ToDateRow = "EOL" Then ToDateRow = R.EOL
            FrDate = .Cells(FrDateRow, Col) ' переписываем FrDate и ToDate
            ToDate = .Cells(ToDateRow, Col)
        End With
        Application.DisplayAlerts = False
            .Sheets(OldRepName).Delete
        Application.DisplayAlerts = True
    End With
    
'---- переписываем FrDate и ToDate в TOCmatch
    With DB_MATCH.Sheets(TOC)
        .Cells(R.iTOC, TOC_FRDATE_COL) = FrDate
        .Cells(R.iTOC, TOC_TODATE_COL) = ToDate
    End With
End Sub
Sub DocReset(DocName As String)
'
' -DocReset(DocName)    - сброс всех шагов, работающих с DocName
'
' 26.10.13
' 27.10.13 - работаем со Стеком Процессов и с контрольной суммой Документа

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
' - StepReset(iStep) - сброс Шага в таблице Процессов - РЕКУРСИЯ!
' 28.8.12
'  9.9.12 - bug fix в сбосе выполненного Шага при загрузке нового Документа
' 13.9.12 - bug fix - не сбрасываем Шаги <*>ProcStart
' 22.10.13 - bug fix - Range колонок 1..3 переписан
' 23.10.13 - сбрасываем не только Шаг iStep, а весь остаток процедуры и ссылки
'            и все Шаги, для которых iStep является "предыдущим" - PrevStep
' 24.10.13 - bug fix - обход обработки ссылки <*>ProcEnd PrevStep

    Dim i As Integer, iProc As Integer
    Dim ThisProc As String, PrevSteps() As String
    Dim PrevS, PrS() As String
    
    With DB_MATCH.Sheets(Process)
        If .Cells(iStep, PROC_STEPDONE_COL) = "" Then Exit Sub
'---- сброс Шагов от iStep до <*>ProcEnd и окраски старта Процедуры "<*>ProcStart"
        For i = 6 To EOL(Process, DB_MATCH)
            If .Cells(i, PROC_STEP_COL) = PROC_START Then iProc = i
            If i >= iStep Then
                .Cells(i, PROC_STEPDONE_COL) = ""
                .Range("A" & i & ":C" & i).Interior.ColorIndex = 0
            End If
            If .Cells(i, PROC_STEP_COL) = PROC_END And i > iStep Then Exit For
        Next i
        .Range("A" & iProc & ":C" & iProc).Interior.ColorIndex = 0
'---- сброс всех Шагов, в которых в PrevStep ссылаются на невыполненные Шаги
        For i = 6 To EOL(Process, DB_MATCH)
            If .Cells(i, PROC_STEP_COL) = PROC_START Then ThisProc = .Cells(i, PROC_NAME_COL)
            If .Cells(i, PROC_STEP_COL) <> PROC_END _
                    And .Cells(i, PROC_COMMENT_COL) <> "" Then
                PrevSteps = Split(.Cells(i, PROC_PREVSTEP_COL), ",")
                For Each PrevS In PrevSteps
                    If InStr(PrevS, "/") = 0 Then
                       If Not IsStepDone(ThisProc, PrevS) Then StepReset i  ' РЕКУРСИЯ!
                    Else
                       PrS = Split(PrevS, "/")
                       If Not IsStepDone(PrS(0), PrS(1)) Then StepReset i   ' РЕКУРСИЯ!
                    End If
                Next PrevS
            End If
        Next i
    End With
End Sub
