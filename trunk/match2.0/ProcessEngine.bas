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
' 14.9.12 П.Л.Храпкин
'
' - ProcStart(Proc)     - запуск Процесса Proc по таблице Process в match.xlsm
' - IsDone(Proc, Step)  - проверка, что шаг Step процесса Proc уже выполнен
' - Exec(Step, iProc)   - вызов Шага Step по строке iProc таблицы Процессов
' - ToStep(Proc,[Step]) - возвращает номер строки таблицы Процессов
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
       
    Dim Code As String, Proc As String
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
Sub StepIn()
'
' - StepIn()    - начало исполнения Шага, т.е. активация и выбор нужных листов
'   1.9.12

    Const FILE_PARAMS = 5   'максимальное количество файлов в Шаге
    
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
' S Adapt(F) - запускает Адаптеры из формы F в match.xlsm
'
' Форма F имеет вид:
'   Шапка   - заголовок колонки. Шапка записывается и форматируется Шагом InsMyCol
'   MyCol   - формулы и раскраска полей от Шапки до Пятки. Если "V" - замена шапки
'   Width   - ширина колонки
'   Columns - номер колонки в активном листе- левом в списке Документов в Процессе
'   Адаптер - строка- вызов Адаптера, обрабатывающего Х = <значение по Columns>
'   Fetch   - строка дополнительных параметров для Адаптера из других Документов
'
' 12.9.12
' 14.9.12 - если Адаптер не нашел значение - оставляем значение по умолчанию

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
' - Adater(Request, X, F_rqst, IsErr) - обрабатывает X в Адаптере "Request"
'    с внешними данными в Документе F_rqst. IsErr=True - ошибка в Адаптере
' 4.9.12
' 6.9.12 - bug fix

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
    Case "CurISO":
        Adapter = CurISO(X)
    Case "CurRate": Adapter = CurRate(X)
    Case "Дата":
        If X = "" Then
            Adapter = ""
        Else
            Adapter = DDMMYYYY(X)
        End If
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
    Dim Doc As String, C1 As Long, C2 As Long, Rng As Range
            
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
        Dim Lit As String
        Const A = 64            ' String("A")-1
        Lit = Chr(C1 + A) & ":" & Chr(C2 + A)
        Set Rng = Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN).Range(Lit)
        
        S = ""
        On Error Resume Next
        S = WorksheetFunction.VLookup(X, Rng, C2 - C1 + 1, False)
        On Error GoTo 0
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
