Attribute VB_Name = "ProcessEngine"
'----------------------------------------------------------------------------------------------------------
' ProcessEngine - процессор Процессов. Состояние и Шаги процессов записаны в листе Process файла match.xlsm
'                   * Процессы Loader'ы запускаются при загрузке в DBs нового Документа.
'                   * Загрузка нового Документа сбрасывает в начальное состояние все Процессы,
'                     связанные с использованием этого Документа
'                   * Процессы Handler'ы и Reporter'ы инициируются из меню формы Главной Панели
'                   * Все Процессы повторновходимы. Если Шаг процесса уже выполнен - он пропускается.
'
'   31.7.12 П.Л.Храпкин
'
' - ProcStart(Proc) - запуск Процесса Proc по таблице Process в match.xlsm

Option Explicit

Sub ProcStart(Proc)
'
' - ProcStart(Proc) - запуск Процесса Proc по таблице Process в match.xlsm
'   31.7.12

    Dim i As Integer
    Dim Step As String          ' Шаг Процесса
    Dim PrevStep As String      ' предыдущий обязательный Шаг
    Dim ProcName As String      ' Имя Процесса
    
    GetRep Process
    With DB_MATCH.Sheets(RepTOC.SheetN)
        For i = 6 To RepTOC.EOL
            Step = .Cells(i, PROC_STEP_COL)
            ProcName = .Cells(i, PROC_NAME_COL)
            If Step = PROC_START And ProcName = Proc Then GoTo ProcStart
        Next i
        ErrMsg FATAL_ERR, "Не найден Процесс " & Proc
        End
ProcStart:
        Dim StepRow As Range
        Dim i0 As Integer   ' строка - начало Процесса
        Dim j As Integer    ' индекс цикла- поиска PrevStep
        i0 = i
        Do While Step <> PROC_END
            i = i + 1
            Step = .Cells(i, PROC_STEP_COL)
            If .Cells(i, PROC_STEPDONE_COL) <> "1" Then
'                StepRow = .EntireRow.Cells(i, 1)
                PrevStep = .Cells(i, PROC_PREVSTEP_COL)
                GetRep .Cells(i, PROC_REP1_COL)
'-- проверяем, подлежит ли выполнению Шаг
                If PrevStep = REP_LOADED _
                        And RepTOC.Made = REP_LOADED Then GoTo ObeyStep
                For j = i0 To i
                    If .Cells(j, PROC_STEP_COL) = PrevStep Then
                        If .Cells(j, PROC_STEPDONE_COL) <> "1" Then GoTo Err
                        GoTo ObeyStep
                    End If
                Next j
                GoTo Err
'-- выполняем Шаг
ObeyStep:
                If .Cells(i, PROC_PAR1_COL + 4) = "" Then
                    Application.Run Step, _
                        .Cells(i, PROC_PAR1_COL), _
                        .Cells(i, PROC_PAR1_COL + 1), _
                        .Cells(i, PROC_PAR1_COL + 2), _
                        .Cells(i, PROC_PAR1_COL + 3)
                ElseIf .Cells(i, PROC_PAR1_COL + 3) = "" Then
                    Application.Run Step, _
                        .Cells(i, PROC_PAR1_COL), _
                        .Cells(i, PROC_PAR1_COL + 1), _
                        .Cells(i, PROC_PAR1_COL + 2)
                ElseIf .Cells(i, PROC_PAR1_COL + 2) = "" Then
                    Application.Run Step, _
                        .Cells(i, PROC_PAR1_COL), _
                        .Cells(i, PROC_PAR1_COL + 1)
                ElseIf .Cells(i, PROC_PAR1_COL + 1) = "" Then
                    Application.Run Step, _
                        .Cells(i, PROC_PAR1_COL)
                ElseIf .Cells(i, PROC_PAR1_COL) = "" Then
                    Application.Run Step
                Else
                    Application.Run Step, _
                        .Cells(i, PROC_PAR1_COL), _
                        .Cells(i, PROC_PAR1_COL + 1), _
                        .Cells(i, PROC_PAR1_COL + 2), _
                        .Cells(i, PROC_PAR1_COL + 3), _
                        .Cells(i, PROC_PAR1_COL + 4)
                End If
                .Cells(i, PROC_STEPDONE_COL) = "1"  ' Done = "1" - Шаг выполнен
            End If
        Loop
        
    End With
Err:
    ErrMsg FATAL_ERR, "Нарушена последовательность шагов процедуры " & ProcName
    End
End Sub
