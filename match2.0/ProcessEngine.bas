Attribute VB_Name = "ProcessEngine"
'----------------------------------------------------------------------------------------------------------
' ProcessEngine - процессор Процессов. Состояние и Шаги процессов записаны в листе Process файла match.xlsm
'                   * Процессы Loader'ы запускаются при загрузке в DBs нового Документа.
'                   * Загрузка нового Документа сбрасывает в начальное состояние все Процессы,
'                     связанные с использованием этого Документа
'                   * Процессы Handler'ы и Reporter'ы инициируются из меню формы Главной Панели
'                   * Все Процессы повторновходимы. Если Шаг процесса уже выполнен - он пропускается.
'
'   2.8.12 П.Л.Храпкин
'
' - ProcStart(Proc) - запуск Процесса Proc по таблице Process в match.xlsm

Option Explicit

Sub ProcStart(Proc)
'
' - ProcStart(Proc) - запуск Процесса Proc по таблице Process в match.xlsm
'   2.8.12

    Dim i As Integer
    Dim Step As String          ' Шаг Процесса
    Dim PrevStep As String      ' предыдущий обязательный Шаг
    Dim ProcName As String      ' Имя Процесса
    
    GetRep Process
    With DB_MATCH.Sheets(Process)
        For i = 6 To RepTOC.EOL
            Step = .Cells(i, PROC_STEP_COL)
            ProcName = .Cells(i, PROC_NAME_COL)
            If Step = PROC_START And ProcName = Proc Then GoTo MyProc
        Next i
        ErrMsg FATAL_ERR, "Не найден Процесс " & Proc
        End
MyProc:
        Dim StepRow As Range
        Do While Step <> PROC_END
            i = i + 1
            Step = .Cells(i, PROC_STEP_COL)
            
'-- Оператор -?-IsDone -- проверка, что используемый Документ обработан
'                      -- если еще не выполнен - его рекурсивный вызов
            If Step = PROC_IFISDONE Then
                ProcName = .Cells(i, PROC_NAME_COL)
                Step = .Cells(i, PROC_NAME_COL + 1)
                If Not IsDone(ProcName, Step) Then ProcStart ProcName
            End If
            
            If .Cells(i, PROC_STEPDONE_COL) <> "1" Then
                PrevStep = .Cells(i, PROC_PREVSTEP_COL)
'-- проверяем, подлежит ли выполнению очередной Шаг по PrevStep
                If PrevStep <> REP_LOADED Then
                    GetRep .Cells(i, PROC_REP1_COL)
                    If RepTOC.Made <> REP_LOADED Then GoTo Err
                End If
                If Not IsDone(Proc, PrevStep) Then GoTo Err

'*************************************
                Exec Step, i        '*  выполняем Шаг
'*************************************
            
            End If
        Loop
        
    End With
    MS "<*> Процесс " & Proc & " завершен!"
    Exit Sub
Err:
    ErrMsg FATAL_ERR, "Нарушена последовательность шагов процедуры " & ProcName
    End
End Sub
Function IsDone(Proc, Step) As Boolean
'
' - IsDone(Proc, Step) - проверка, что шаг Step процесса Proc выполнен
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
        ErrMsg FATAL_ERR, "IsDone: Обращение к несуществующему Процессу " & Proc
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
        ErrMsg FATAL_ERR, "IsDone: Обращение к несуществующему Шагу " & Step _
            & " Процесса " & Proc
    End With
End Function
Sub Exec(Step, iProc)
'
' - Exec(Step, iProc) - вызов Шага Step по строке iProc таблицы Процессов
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
        .Cells(iProc, PROC_STEPDONE_COL) = "1"  ' Done = "1" - Шаг выполнен
        .Cells(iProc, PROC_TIME_COL) = Now
        .Cells(1, 1) = Now
    End With
End Sub
Sub testRunProc()
    RunProc "REP_SF_LOAD"
End Sub
Sub RunProc(Proc)
'
' - RunProc(Proc)   - запуск Процесса Proc
'   31.7.12
'''    GetRep Process
    Application.Run "'" & DirDBs & F_MATCH & "'!ProcStart", Proc
End Sub
