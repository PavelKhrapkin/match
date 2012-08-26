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
'   26.8.12 П.Л.Храпкин
'
' - ProcStart(Proc)     - запуск Процесса Proc по таблице Process в match.xlsm
' - IsDone(Proc, Step)  - проверка, что шаг Step процесса Proc уже выполнен
' - Exec(Step, iProc)   - вызов Шага Step по строке iProc таблицы Процессов
' - ToStep(Proc,[Step]) - возвращает номер строки таблицы Процессов

Option Explicit

Const TRACE_STEP = "Trace"  ' специальный Шаг Trace для трассирования и отладки
Public TraceStep As Boolean
Public TraceStop As Boolean

Sub ProcStart(Proc As String)
'
' - ProcStart(Proc) - запуск Процесса Proc по таблице Process в match.xlsm
'   7.8.12
'  26.8.12 - окраска выполненного Процесса

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
                    
                .Cells(1, PROCESS_NAME_COL) = Proc      'имя Процесса
                .Cells(1, STEP_NAME_COL) = Step         'имя Шага
                
'*************************************
                Exec Step, i        '*  выполняем Шаг
'*************************************
            
            End If
        Loop
        .Cells(1, PROCESS_NAME_COL) = "": .Cells(1, STEP_NAME_COL) = ""
        .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 35
    End With
    MS "<*> Процесс " & Proc & " завершен!"
    Exit Sub
Err:
    ErrMsg FATAL_ERR, "Нарушена последовательность шагов процедуры " & Proc
    End
End Sub
Function IsDone(ByVal Proc As String, ByVal Step As String) As Boolean
'
' - IsDone(Proc, Step) - проверка, что шаг Step процесса уже Proc выполнен
'   7.8.12
'  16.8.12 - bug fix про PrevStep без запятой давал ошибку

    Dim i As Integer
    Dim iStep As Long
    Dim S() As String   '=части требований PrevStep, разделенные ","
    Dim X() As String   '=каждая часть может быть вида <Proc>/<Step>
    Dim Rep As String, Done As String
    
    Proc = Trim(Proc): Step = Trim(Step)
    
    If Step = REP_LOADED Then
        i = ToStep(Proc)
        Rep = DB_MATCH.Sheets(Process).Cells(i, PROC_REP1_COL)
        GetRep Rep
        If RepTOC.Made <> REP_LOADED Then
            Dim Msg As String
            ErrMsg FATAL_ERR, "IsDone: Не 'Loaded' файл для Процесса " _
                & Proc & " на Шаге " & Step & vbCrLf & vbCrLf _
                & "Отчет " & RepTOC.Name & " надо загрузить заново!"
            Stop
            End
        Else
            If TraceStep Then MS "Отчет " & Rep & " действительно 'Loaded'"
            IsDone = True
            Exit Function
        End If
    Else
        S = split(Trim(Step), ",")
        For i = LBound(S) To UBound(S)
            If InStr(S(i), "/") <> 0 Then
                X = split(S(i), "/")
                If Proc = X(0) Then ErrMsg FATAL_ERR, "Бесконечная рекурсия в PrevStep!!"
                If Not IsDone(X(0), X(1)) Then ProcStart X(0)
            Else
                iStep = ToStep(Proc, S(i))
                If DB_MATCH.Sheets(Process).Cells(iStep, PROC_STEPDONE_COL) <> "" Then
                    IsDone = True
                    Exit Function
                End If
                ProcStart Proc
            End If
        Next i
        IsDone = True
        Exit Function
    End If
End Function
Sub Exec(Step, iProc)
'
' - Exec(Step, iProc) - вызов Шага Step по строке iProc таблицы Процессов
'   7.8.12
'  26.8.12 - окраска строки в Process для успешно выполненного Шага
       
    Dim Code As String
    Dim File As String
    Dim R As TOCmatch       '= обрабатываемый Документ - отчет
            
    If Step = PROC_END Or Step = "" Then Exit Sub
    
    With DB_MATCH.Sheets(Process)
'-- Trace - специальный Шаг для запуска трассирования и отладки Шагов
        If Not TraceStep Then TraceStep = False
        If Step = TRACE_STEP Then
            TraceStep = True
            TraceStop = False
            If .Cells(iProc, PROC_PAR1_COL) = 1 Then TraceStop = True
            Exit Sub
        End If

'*********** вызов подпрограммы - Шага ***********************
'// в будущем, когда от Loader'ов перейдем к Handler'ам
'// имя "рабочего" отчета можно будет брать из строки - названия Процедуры

        R = GetRep(.Cells(iProc, PROC_REP1_COL))
            '-- Select лист, с которым будем работать
        If R.Name <> "" Then Workbooks(R.RepFile).Sheets(R.SheetN).Select
        Code = Step
        PublicStepName = Step
'        File = .Cells(iProc, PROC_STEPFILE_COL)
        File = F_MATCH  '!!!!!!!!!!!!!!!!!!!!!!!!
        If File <> "" Then Code = "'" & DirDBs & File & "'!" & Step
        
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
        Application.StatusBar = False
        .Cells(iProc, PROC_STEPDONE_COL) = "1"  ' Done = "1" - Шаг выполнен
        .Cells(iProc, PROC_TIME_COL) = Now
        .Range(Cells(iProc, 1), Cells(iProc, 3)).Interior.ColorIndex = 35
        .Cells(1, STEP_NAME_COL) = ""
        .Cells(1, 1) = Now
        R.Made = Step
        R.Dat = Now
        RepTOC = R
        WrTOC
    End With
End Sub
Function ToStep(Proc, Optional Step As String = "") As Integer
'
' - ToStep(Proc, [Step]) - возвращает номер строки таблицы Процессов
'   7.8.12
    
    Dim P As TOCmatch           'строка таблицы Процессов в виде TOCmatch
    Dim StepName As String      '=Имя текущего Шага
    Dim ProcName As String      '=Имя текущего Процесса
    Dim i As Integer
    
    P = GetRep(Process)
    
    With DB_MATCH.Sheets(Process)
        For i = 6 To RepTOC.EOL
            ProcName = .Cells(i, PROC_NAME_COL)
            StepName = .Cells(i, PROC_STEP_COL)
            If StepName = PROC_START And ProcName = Proc Then GoTo MyProc
        Next i
        ErrMsg FATAL_ERR, "ToProc: Не найден Процесс " & Proc
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
Sub testRunProc()   'Ctrl/W
Attribute testRunProc.VB_ProcData.VB_Invoke_Func = "W\n14"
'    RunProc "REP_1C_P_LOAD"
'    RunProc "REP_SF_LOAD"
'    RunProc "HANDL_Payment"
'    RunProc "HANDL_NewContract"
    RunProc "REPORT_ADSKquantity"
End Sub
Sub RunProc(Proc)
'
' - RunProc(Proc)   - запуск Процесса Proc
'   31.7.12
'''    GetRep Process
    Application.Run "'" & DirDBs & F_MATCH & "'!ProcStart", Proc
End Sub
