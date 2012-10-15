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
' 14.10.12 П.Л.Храпкин
'
' - ProcStart(Proc)     - запуск Процесса Proc по таблице Process в match.xlsm
' - IsDone(Proc, Step)  - проверка, что шаг Step процесса Proc уже выполнен
' - Exec(Step, iProc)   - вызов Шага Step по строке iProc таблицы Процессов
' - ToStep(Proc,[Step]) - возвращает номер строки таблицы Процессов
' - ToProcEnd(iProc)    - позиционирование на <*>ProcEnd таблицы Процессов
' S ProcReset(Proc)     - сброс и новый запуск Процесса Proc
' - StepIn()            - начало исполнения Шага, т.е. активация нужных листов
' S Adapt(F) - запускает Адаптеры из формы F
' - Adater(Request, X, F_rqst, IsErr) - обрабатывает X в Адаптере "Request"
'        с внешними данными в Документе F_rqst. IsErr=True - ошибка в Адаптере
' - FetchDoc(F_rqst, X, IsErr) - извлечение данных из стороннего Документа
'                   по запросу F_rqst для значения поля X. IsErr=True - ошибка

Option Explicit

Const TRACE_STEP = "Trace"  ' специальный Шаг Trace для трассирования и отладки
Public TraceStep As Boolean
Public TraceStop As Boolean

'----- работа с Адаптерами ---------------
Const WP_CONTEXT_LINE = 8, WP_CONTEXT_COL = 4   ' ячейка передачи iLine

Const EXT_PAR = "ExtPar"    ' текст в Шаблоне - признак передачи параметра Х

Const PTRN_VALUE = 2 ' смещение строки - значения - Value в Шаблоне
Const PTRN_WIDTH = 3 ' смещение строки - ширина колонок в Шаблоне
Const PTRN_COLS = 4  ' смещение строки номеров колонок в Шаблоне
Const PTRN_ADAPT = 5 ' смещение строки вызова Адаптеров в Шаблоне
Const PTRN_FETCH = 6 ' смещение строки вызова Fetch - извлечения из Док-в в Шаблоне
Const PTRN_LNS = 6   ' кол-во строк в Шаблоне по каждой группе строк на экране

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
        .Cells(1, PROCESS_NAME_COL) = "": .Cells(1, STEP_NAME_COL) = ""
        .Range(Cells(i + 1, 1), Cells(i + 1, 2)).Interior.ColorIndex = 35
    End With
''    MS "<*> Процесс " & Proc & " завершен!"
    Exit Sub
Err:
    ErrMsg FATAL_ERR, "Нарушена последовательность шагов процедуры " & Proc
    End
End Sub
Function IsDone(ByVal Proc As String, ByVal Step As String) As Boolean
'
' - IsDone(Proc, Step) - проверка, что шаг Step процесса уже Proc выполнен,
'                        а если не выполнен - запуск исполняющей его Процедуры
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
                ProcStart Proc  'здесь - только по PrevStep
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
            If .Cells(iProc, PROC_PAR1_COL) = 1 Then TraceStop = True
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
Sub StepOut(Step, iProc)
'
' - StepOut()   - завершение выполнения Шага с записью в TOCmatch
'   8/10/12

    Dim Proc As String, R As TOCmatch
    
    With DB_MATCH.Sheets(Process)
        Application.StatusBar = False
        .Activate
        .Cells(iProc, PROC_STEPDONE_COL) = "1"  ' Done = "1" - Шаг выполнен
        .Cells(iProc, PROC_TIME_COL) = Now
        .Range(Cells(iProc, 1), Cells(iProc, 3)).Interior.ColorIndex = 35
        .Cells(1, STEP_NAME_COL) = ""
        .Cells(1, 1) = Now
        Proc = .Cells(1, PROCESS_NAME_COL)              'имя Процесса
        If Proc = "" Then Exit Sub
        R = GetRep(.Cells(ToStep(Proc), PROC_REP1_COL)) 'обрабатываемый Документ
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
Sub ProcReset(Proc As String)
'
' S ProcReset(Proc) - сброс и новый запуск Процесса Proc
' 1.10.12

    Dim i As Long
    
    GetRep Process
    With DB_MATCH.Sheets(Process)
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
' S CheckProc0(NewProcResult)   - проверка, что вспомогательный Процесс не нашел
'                                 новых "автоматических" записей в SF
' 1/10/12

    If NewProcResult <> "0" Then
        ErrMsg FATAL_ERR, PublicProcName & ": CheckProc0> в результате не '0'"
        End
    End If
End Sub
Sub WrNewSheet(SheetNew, SheetDB, DB_Line, Optional ExtPar As String)
'
' - WrNewSheet(SheetNew, SheetDB, DB_Line) - записывает новый рекорд в лист SheetNew
'                                            из строки DB_Line листа SheetDB
'   * Имя и Параметры для обработки передаются в Адаптер в виде текстовых строк.
'     Эти строки хранятся в Range с именем "HDR_" & SheetNew в Forms или Headers
'   * Обращение к Адаптеру имеет вид <ИмяАдаптера>/<Пар1>,<Пар2>...
'   * В строке формы под Адаптером можно указать параметры во внешних Документах
' 6.9.2012

    Dim P As Range
    Dim iNewLine As Long    '= номер строки в SheetNew
    Dim i As Long
    Dim X As String         '= обрабатываемое значение в SheetDB
    Dim sX As String        'поле в строке PTRN_COLS Шаблона
    Dim Y As String         '= результат работы Адаптера
    Dim IsErr As Boolean    '=True если Адаптер обнаружил ошибку
    
    iNewLine = EOL(SheetNew, DB_MATCH) + 1

    With DB_MATCH.Sheets(SheetNew)
        Set P = DB_MATCH.Sheets(Header).Range("HDR_" & SheetNew)
        For i = 1 To P.Columns.Count
            sX = P.Cells(PTRN_COLS, i)
            If sX = EXT_PAR Then
                X = ExtPar
            Else
                X = SheetDB.Cells(DB_Line, CLng(sX))
            End If
            
            Y = Adapter(P.Cells(PTRN_ADAPT, i), X, P.Cells(PTRN_FETCH, i), IsErr)
            
            If IsErr Then
                .Rows(iNewLine).Delete
                Exit For
            Else
                .Cells(iNewLine, i) = Y
            End If
        Next i
    End With
End Sub
Sub xAdapt(F As String, iLine As Long)
'
' - xAdapt(F, iLine) - запускает Адаптеры из формы F, обрабатывая данные с экрана
'                      по строке номер iLine в ActiveSheet
'   13.10.12

    Const WP_PROTOTYPE = "WP_Prototype"
    Const PTRN_SELECT = "Select"

    Dim R As TOCmatch                           ' обрабатываемый Документ
    Dim iRow As Integer, iCol As Integer        ' строка и колонка Шаблона F
    Dim PtrnType As String                      ' поле Тип Шаблона
    Dim sX() As String                          ' строка - выражение с аргументом Х
    Dim iX As Long                              ' номер колонки - значение в строке PTRN_COLS
    Dim X As String                             ' параметр Адаптера
    Dim Rqst As String                          ' строка - обращение к Адаптеру
    Dim F_rqst As String                        '
    Dim Y As String
    Dim IsErr As Boolean
    Dim iSelect As Long, WP_Row As Long
        
    Set DB_TMP = FileOpen(F_TMP)
    Application.DisplayAlerts = False
    On Error Resume Next
    DB_TMP.Sheets(WP).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    DB_MATCH.Activate
    DB_MATCH.Sheets("WP_Prototype").Copy Before:=DB_TMP.Sheets(1)
    ActiveSheet.Name = WP
    
    With DB_TMP.Sheets(WP)
        .Rows("1:4").Delete
        .Tab.Color = rgbBlue
        
        .Cells(1, 1) = "'" & DirDBs & F_MATCH & "'!xAdapt_Continue"
        .Cells(WP_CONTEXT_LINE, WP_CONTEXT_COL) = iLine
        For iRow = 1 To .UsedRange.Rows.Count Step PTRN_LNS
            PtrnType = .Cells(iRow, 2)
            
            R.EOL = -1                      ' инициализация EOL на случай Select
            If PtrnType <> "Кнопки" Then
                R = GetRep(.Cells(iRow, 1))
                Workbooks(R.RepFile).Sheets(R.SheetN).Activate
            End If
        '--- SelectLoop
            iSelect = 2
            WP_Row = iRow - 1 + PTRN_VALUE
            If PtrnType = PTRN_SELECT Then WP_Row = iRow + PTRN_LNS
            Do
                For iCol = 5 To .UsedRange.Columns.Count
                    sX = split(.Cells(iRow - 1 + PTRN_COLS, iCol), "/")
                    If UBound(sX) > 0 Then iX = sX(0)
                    iX = .Cells(iRow - 1 + PTRN_COLS, iCol)
                    If iX > 0 Then
    
    '''Const PTRN_TYPE_BUTTON = "Кнопки"   'Кнопки, управляющие работой WP
    '''Const PTRN_TYPE_ILINE = "iLine" 'Аргументы X для Адаптеров вычисляются по iLine
    '''Const PTRN_TYPE_PTRN = "Шаблон" 'Аргументы Х для Адаптеров беруться из самого Шаблона
    '''Const PTRN_TYPE_SELECT_OPP = "SelectOpp" ' Аргументы Х и выбор проводится в SelectOpp
                
                        Select Case PtrnType
                        Case "Кнопки":
                            '''' не написан
                        Case "iLine":
                            X = ActiveSheet.Cells(iLine, iX)
                        Case "Шаблон":
                            X = .Cells(iRow - 1 + PTRN_VALUE, iX)
                        Case PTRN_SELECT:
                            X = ActiveSheet.Cells(CLng(Cells(WP_Row, 5)), iX)
                         Case Else:
                            ErrMsg FATAL_ERR, "xAdapt> Странный тип Шаблона " & PtrnType
                        End Select
                        Rqst = .Cells(iRow - 1 + PTRN_ADAPT, iCol)
                        F_rqst = .Cells(iRow - 1 + PTRN_FETCH, iCol)
                        
                        Y = Adapter(Rqst, X, F_rqst, IsErr, R.EOL)
                        
                        If Not IsErr Then
                            .Cells(WP_Row, iCol) = Y
                            If UBound(sX) > 0 Then
                                Select Case sX(1)
                                Case "CopyVal":
                                    .Cells(iRow - 1 + PTRN_COLS, iCol).Copy .Cells(WP_Row, iCol)
                                Case "":
                                Case Else
                                    ErrMsg FATAL_ERR, "xAdapt> неправильный Шаблон в [" _
                                        & iRow - 1 + PTRN_COLS & ", " & iCol & "]"
                                    End
                                End Select
                            End If
                        End If
                    ElseIf iX < 0 Then
                        Exit For
                    End If
                Next iCol
                If PtrnType = PTRN_SELECT Then
                    iSelect = .Cells(WP_Row, iCol)
                    .Cells(iRow - 1 + PTRN_VALUE, 4) = iSelect
                    .Rows(iRow - 1 + PTRN_COLS).Hidden = True
                End If
                WP_Row = WP_Row + 1
            Loop While PtrnType = PTRN_SELECT And iSelect < R.EOL
                
            .Rows(iRow - 1 + PTRN_COLS).Hidden = True
            .Rows(iRow - 1 + PTRN_ADAPT).Hidden = True
            .Rows(iRow - 1 + PTRN_WIDTH).Hidden = True
            .Rows(iRow - 1 + PTRN_FETCH).Hidden = True
        Next iRow
    End With
    
'=====  СОХРАНЕНИЕ КОНТЕКСТА ====================
    
'''''''''''''''''''''''''''''''''''
    End '''  остановка VBA ''''''''
'''''''''''''''''''''''''''''''''''
End Sub
Sub xAdapt_Continue(Button As String)
'
' * xAdapt_Continue(Button) - продолжение работы Adapt после нажатия кнопки Button
'                             Сюда передается управления из WP_Select_Button.
' 8/10/12

    Dim Step As String, iLine As Long
        
    GetRep (Process)
    DB_MATCH.Activate
    iLine = ActiveSheet.Cells(WP_CONTEXT_LINE, WP_CONTEXT_COL)
    
    Select Case Button
    Case "STOP":
        Step = DB_MATCH.Sheets(Process).Cells(STEP_NAME_COL, 1)
        StepOut Step, ToStep(DB_MATCH.Sheets(Process).Cells(1, PROCESS_NAME_COL), Step)
    Case "->":
        WP_PdOpp WP, iLine
    Case "NewOpp":
    Case "Проект":
    End Select
End Sub
Function AdaptLine(XXX, FF As Range, F_Row As Integer) As Boolean
'
'
'

    With FF
        For iCol = 1 To .Columns.Count
            iX = FF(F_Row + PTRN_COLS, iCol)
            If iX > 0 Then
                X = Sht.Cells(i, iX)
                Rqst = FF.Cells(5, Col)
                F_rqst = FF.Cells(6, Col)
                
                Y = Adapter(Rqst, X, F_rqst, IsErr)
                
                If Not IsErr Then .Cells(i, Col) = Y
            ElseIf iX < 0 Then
                Exit For
            End If
        Next iCol
    End With
End Function
Sub Adapt(F As String)
'
' S Adapt(F) - запускает Адаптеры из формы F, осуществляя проход по Документу
'
' Форма F имеет вид:
'   Шапка   - заголовок колонки. Шапка записывается и форматируется Шагом InsMyCol
'   MyCol   - формулы и раскраска полей от Шапки до Пятки. Если "V" - замена шапки
'   Width   - ширина колонки
'   Columns - номер колонки в активном листе- левом в списке Документов в Процессе
'       - пустое поле Columns - целевое поле остается без изменения
'       <0  - выход из цикла по колонкам
'   Адаптер - строка- вызов Адаптера, обрабатывающего Х = <значение по Columns>
'   Fetch   - строка дополнительных параметров для Адаптера из других Документов
'
' 12.9.12
' 14.9.12 - если Адаптер не нашел значение - оставляем значение по умолчанию
' 26.9.12 - обработка пустых и отрицательных значений Columns

    StepIn
    
    Dim FF As Range     '= Форма F
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
                iX = FF(PTRN_COLS, Col)
                If iX > 0 Then
                    X = .Cells(i, iX)
                    Rqst = FF.Cells(PTRN_ADAPT, Col)
                    F_rqst = FF.Cells(PTRN_FETCH, Col)
                    
                    Y = Adapter(Rqst, X, F_rqst, IsErr)
                    
                    If Not IsErr Then .Cells(i, Col) = Y
                ElseIf iX < 0 Then
                    Exit For
                End If
            Next Col
        Next i
    End With
End Sub
Function Adapter(Request, ByVal X, F_rqst, IsErr, Optional EOL_Doc) As String
'
' - Adater(Request, X, F_rqst, IsErr) - обрабатывает X в Адаптере "Request"
'    с внешними данными в Документе F_rqst. IsErr=True - ошибка в Адаптере
' 4.9.12
' 6.9.12 - bug fix
'25.9.12 - Dec(CurRate)
' 3.10.12 - Адаптер GetCol с синтаксисом ' GetCol/1C.xlsx,Платежи,5/SF:2:11
'12.10.12 - Адаптер GoodType(X)
'14.10.12 - Адаптер OppFilter для Шаблона типа Select

    Dim FF() As String, Tmp() As String, Cols() As String
    Dim Doc As String, C1 As Long, C2 As Long, Rng As Range
    Dim F() As String
    Dim i As Long, Par() As String
    
    IsErr = False
    X = Trim(X)
    
'--- разбор строки Адаптера вида <Имя>/C1,C2,C3...
    Dim AdapterName As String
    AdapterName = ""
    If Request <> "" Then
        Tmp = split(Request, "/")
        AdapterName = Tmp(0)
        If InStr(Request, "/") <> 0 Then Par = split(Tmp(1), ",")
    End If

'========== препроцессинг Адаптера =========
    Select Case AdapterName
    Case "MainContract":
        X = Trim(Replace(X, "Договор", ""))
    Case "<>0":
        If X = "0" Then X = ""
    End Select
    
'--- FETCH разбор строки параметров из Документов вида <Doc1>/C1:C2,<Doc2>/C1:C2,...
    If F_rqst <> "" And X <> "" Then
        
        FF = split(F_rqst, ",")
        For i = LBound(FF) To UBound(FF)
            X = FetchDoc(FF(i), X, IsErr)
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' сейчас используется только один указатель на извлекаемую из Doc величину.
' В дальнейшем надо использовать массив x(1 to 5) и обращаться к Fetch несколько раз
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Next i
    End If

'******* выполнение Адаптера с параметрами Par ******
    Select Case AdapterName
    Case "", "MainContract": Adapter = X
    Case "Мы", "Продавец_в_SF", "Vendor":
        On Error GoTo AdapterFailure
        Adapter = WorksheetFunction.VLookup(X, DB_MATCH.Sheets("We").Range(AdapterName), Par(0), False)
        On Error GoTo 0
    Case "Dec": Adapter = Dec(X)
    Case "GetCol":
        If X = "" Then
            Adapter = ""
        Else                ' GetCol/1C.xlsx,Платежи,5 [/SF/2:11]
            Adapter = Workbooks(Par(0)).Sheets(Par(1)).Cells(CLng(X), CLng(Par(2)))
            If UBound(Tmp) > 1 Then
                Adapter = FetchDoc(Tmp(2) & "/" & Tmp(3), Adapter, IsErr)
            End If
        End If
    Case "GoodType":
        Adapter = GoodType(X)
    Case "CurISO":
        Adapter = CurISO(X)
    Case "CurRate": Adapter = Dec(CurRate(X))
    Case "Дата":
        If X = "" Then
            Adapter = ""
        Else
            Adapter = DDMMYYYY(X)
        End If
    Case "OppFilter":
        With DB_TMP.Sheets(WP)
            Const SEL_REF = 20
            Dim b As Long, A(0 To 6) As Long
            b = .Cells(SEL_REF + 2, 4)
            For i = 0 To UBound(A)
                A(i) = CLng(Par(i))
            Next i
            For i = .Cells(SEL_REF, 4) + 1 To EOL_Doc
                If OppFilter(i, .Cells(b, A(0)), .Cells(b, A(1)), _
                        .Cells(b, A(2)), .Cells(b, A(3)), .Cells(b, A(4)), _
                        .Cells(b, A(5)), .Cells(b, A(6))) Then
                    Adapter = i
                    Exit For
                End If
            Next i
            ActiveSheet.Cells(i, 5) = X
        End With
    Case Else
        ErrMsg FATAL_ERR, "Adapter> Не существует " & AdapterName
    End Select
    Exit Function
    
AdapterFailure:
    ErrMsg WARNING, "Адаптер " & AdapterName & "(" & X & ") не получил данных"
    IsErr = True
End Function
Function FetchDoc(F_rqst, X, IsErr) As String
'
' - FetchDoc(F_rqst, X, IsErr) - извлечение данных из стороннего Документа
'                   по запросу F_rqst для значения поля X. IsErr=True - ошибка
'
' * F_rqst имеет вид <Doc>/C1[:C2][/W]
' * <Doc>   - имя Документа, окуда извлекаются данные
' *   /     - отделяет группы параметров. Возможно несколько групп.
' *   :     - разделяет параметры внутри группы
' *             Первая группа - собственно параметры для извлечения С1[:С2]
' *  C1                 если есть только С1 - извлекается поле номер С1
' * C1:C2               если С1:С2 - Lookup по С1 -> из C2 с Range по Doc
' *             Вторая группа - параметры обработки ошибок Fetch - /W или /0
' *  /W             - WARNING в Log, оставлять IsErr=False, если извлечено ""
' *  /0             - "" вполне допустимо (например, область в адресе)
' *  /D             - Default - "" допустим, но IsErr=True для значения по умолчанию
' *             Вторая группа отсутствует - выводить Log и IsErr = True
'
' 5.9.12
' 14.9.12 - работает /D для второй группы - "по умолчанию"

    FetchDoc = ""
    If F_rqst = "" Or X = "" Then GoTo ErrExit
        
    Dim Tmp() As String, Cols() As String, S As String
    Dim Doc As String, C1 As Long, C2 As Long, Rng As Range, N As Long
            
    Tmp = split(F_rqst, "/")
    Doc = Tmp(0)
    Cols = split(Tmp(1), ":")
    C1 = Cols(0)
    
    Dim Rdoc As TOCmatch, W As Workbook
    Rdoc = GetRep(Doc)
    
    If UBound(Cols) < 1 Then
'--- ситуация С1 - в группе один параметр - извлекаем значение по индексу
        Dim Indx As Long
        Indx = X
'!!!!!!!!!!!!!!!!!!!!!!!!!!!
' сейчас Indx=Х - это просто число, но в дальнейшем тут надо Split
'!!!!!!!!!!!!!!!!!!!!!!!!!!!
        If Indx <= 0 Then
            ErrMsg WARNING, "FetchDoc: " & Doc & "(" & Indx & "," & C1 _
                & ") - неправильный номер строки"
            GoTo ErrExit
        End If
        S = Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN).Cells(Indx, C1)
    Else
'--- ситуация С1:C2 - в группе 2 параметра - извлекаем значение по Lookup
        C2 = Cols(1)
        S = ""
        N = CSmatchSht(X, C1, Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN))
        If N <> 0 Then S = Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN).Cells(N, C2)
    End If
'--- обработка группы 2 -- если S=""
    If S = "" Then
        If UBound(Tmp) >= 2 Then
            If Tmp(2) = "W" Then
                ErrMsg WARNING, "Адаптер> ссылка " & F_rqst _
                    & "(" & X & ") не работает, результат <пусто>"
            End If
            If Tmp(2) <> "0" Then GoTo ErrExit
        Else
            ErrMsg WARNING, "Адаптер> ссылка " & F_rqst _
               & "(" & X & ") не работает, результат <пусто>"
            GoTo ErrExit
        End If
    Else
        FetchDoc = S
    End If
    
OK_Exit:    IsErr = False
    Exit Function
ErrExit:    IsErr = True

End Function
