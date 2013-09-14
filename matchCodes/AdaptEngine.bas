Attribute VB_Name = "AdaptEngine"
'---------------------------------------------------------------------------------------
' AdaptEngine - процессор Адаптеров.
'       * Адаптер - синтаксическая единица, предназначенные для обработки поля Документа
'       * Описания Адаптеров хранятся в Шаблоне, в том числе:
'           1.(кол.1) Основной обрабатываемый Документ - он становится ActiveSheet
'           2.(кол.2) тип Шаблона (используется в Шаблонах WP)
'               2.1 Кнопки (New, ->, Stop)
'               2.2 iLine - ссылка на строку номер iLine, получаемую как параметр WP_Adapt
'               2.3 Шаблон - наиболее распостраненный тип Шаблона общего назначения
'               2.4 Select - содержит Адаптер OppSelect, выводящий НЕСКОЛЬКО строк-записей
'       * Первая строка Шаблона содержит "шапку" - заголовки колонок
'       * Шаг, вызывающий Адаптеры Шаблона, делает цикл по строкам основного Документа
'       * Вторая строка - Value - содержит формулы и результаты Y, вычисляемые Адаптерами
'           - Шаблон Select выводит записи, выбранные Адаптером OppFilter ниже Шаблона
'           - такие строки, выбранные в Select сами содержат КНОПКИ - действия по записям
'           - некоторые поля Шаблона в строке Value содержат ГИПЕРССЫЛКИ для вызова SF
'       * Третья строка - Width - ширину выводимой колонки и другие форматные атрибуты
'           - число - ширина в условных единицах - около 1 мм
'           - признаки Dbl,..
'           - для Шаблона Select ширина колонки определяется только первой строкой Width
'       * Четвертая строка - Columns или Х - определение аргумента Адаптера
'           - разбор синтаксиса проводится в X_Parse
'           - #6 - обращение к колонку 6 в Value самого Шаблона, а не ActiveSheet
'           - !6 - тоже обращение к Шаблону, но в Select Х берется из выбранной строки
'       * Пятая строка - Адаптер - происходит вычисление Y = Adapter (X)
'       * Шестая строка - Fetch - перед обращением к Адаптеру проводится его извлечение
'         из других Документов обращением вида X = SFD/18:2, то есть X из Четвертой строки
'         используется для Lookup в Документе SFD: его значение находится в строке 18, а
'         значение в колонке 2 найденной строки передается Адаптеру как входной аргумент.
'
' 14.09.13 П.Л.Храпкин, А.Пасс
'   История модуля:
' 11.11.12 - выделение AdaptEngine из ProcessEngine
'  7.12.12 - введены форматы вывода "Dbl", "Txt", "Date" в строке "width" в sub WP_Adapt
'  8.12.12 - введен прoизвольный формат в строке width
' 14.12.12 - добавлена обработка формата в строке PTRN_WIDTH (WrNewSheet)
' 17.12.12 - добавлен тест целого формата в testfmtCell()
' 19.12.12 - изменен разделитель троек в Dbl в testfmtCell()
' 3.1.13   - введено профилирование Adapt
' 5.1.13   - Merged with Google source
' 8.4.13   - форма ссылки #2 в строке Адаптера "Columns" значит "колонка 2 текушей строки
'                                                                выходного файла"
' 1.5.13   - передача опционального массива параметров в WrNewSheet
'
' - WrNewSheet(SheetNew, SheetDB, DB_Line[,IdOpp]) - записывает новый рекорд
'                               в лист SheetNew из строки DB_Line листа SheetDB
'...........................
' - WP_Adapt(F, iLine)    - выводит на экран WP по Платежу в iLine, подготавливая диалог,
'                           где F - прототип выводимой формы документа (WP_Prototype)
' * WP_Adapt_Continue(Button) - продолжение работы WP_Adapt после нажатия кнопки Button
'...........................
' S Adapt(F) - запускает Адаптеры из Шаблона F
' - Adater(Request, X, F_rqst, IsErr) - обрабатывает X в Адаптере "Request"
'        с внешними данными в Документе F_rqst. IsErr=True - ошибка в Адаптере
' - X_Parse(iRow, iCol, PutToRow, PutToCol)    - разбор строки Х - параметра Адаптера
' - FetchDoc(F_rqst, X, IsErr) - извлечение данных из стороннего Документа
'                   по запросу F_rqst для значения поля X. IsErr=True - ошибка

Option Explicit

'========== Константы и общие переменные Адаптеров ==================
Const WP_CONTEXT_LINE = 8, WP_CONTEXT_COL = 4   ' ячейка передачи iLine
Const WP_PAYMENT_LINE = 8                       ' строка Платежа в WP

Const EXT_PAR = "ExtPar"    ' текст в Шаблоне - признак передачи параметра Х

Const PTRN_HDR = 1   ' смещение строки - значения шапки в Шаблоне
Const PTRN_VALUE = 2 ' смещение строки - значения - Value в Шаблоне
Const PTRN_WIDTH = 3 ' смещение строки - ширина колонок в Шаблоне
Const PTRN_COLS = 4  ' смещение строки номеров колонок в Шаблоне
Const PTRN_ADAPT = 5 ' смещение строки вызова Адаптеров в Шаблоне
Const PTRN_FETCH = 6 ' смещение строки вызова Fetch - извлечения из Док-в в Шаблоне
Const PTRN_LNS = 6   ' кол-во строк в Шаблоне по каждой группе строк на экране

Const FOR_PROCESS = "ForProcess"    'колонка в Шаблоне для проверки
                                    '.. в Адаптере имени Процесса
Const PTRN_SELECT = "Select"
Const OPP_BALKY = "Расходные материалы и ЗИП"
Sub WrNewSheet(SheetNew As String, SheetDB As String, DB_Line As Long, _
    Optional ExtPar As String)
'
' - WrNewSheet(SheetNew, SheetDB, DB_Line[,IdOpp]) - записывает новый рекорд
'                               в лист SheetNew из строки DB_Line листа SheetDB
'   * Используется Шаблон с Адаптерами "HDR_" & SheetNew.
'     Шаблоны обычно храняться в Headers или WP_Prototype
'   * Обращение к Адаптеру имеет вид <ИмяАдаптера>/<Пар1>,<Пар2>...
'   * Если в Шаблоне в строке PTRN_COLS указано "ExtPar", необходимо указать
'                                              параметр ExtPar = IdOpp
' 6.9.2012
' 26.10.12 - обработка "голубых" листов в DB_TMP
' 27.10.12 - использование TOCmatch для "голубых" листов
' 28.10.12 - параметр SheetDB - передается в виде String
' 14.12.12 - добавлена обработка формата в строке PTRN_WIDTH
' 28.01.13 - width в setColWidth теперь массив: ширина/формат
' 8.4.13   - форма ссылки #2 в строке Адаптера "Columns" значит
'            "колонка 2 текушей строки выходного файла"
' 1.5.13   - передача массива доп.параметров для Columns в виде ExtArr/3,
'            то есть третий элемент переданного массива доп.параметров
' 28.08.13 - WrTOC SheetNew
'  6.09.13 - диагностическое сообщение при IsErr=True
'  8.09.13 - использование X_Pars

''    Dim Rnew As TOCmatch, Rdoc As TOCmatch
    Dim P As Range
    Dim i As Long
    Dim X As String         '= обрабатываемое значение в SheetDB
''    Dim sX As String        'поле в строке PTRN_COLS Шаблона
''    Dim sXarr() As String   'номер элемента в массиве ExtPar, напр., ExtPar/2
    Dim Y As String         '= результат работы Адаптера
    Dim IsErr As Boolean    '=True если Адаптер обнаружил ошибку
    Dim NewEOL As Long      '=EOL(SheetNew)
    Dim Width() As String
    
    GetRep (SheetNew)
''    Rnew = GetRep(SheetNew)
''    Rnew.EOL = EOL(Rnew.SheetN, DB_TMP) + 1
''    Rnew.Made = "WrNewSheet"
''    Rdoc = GetRep(SheetDB)
      
    With DB_TMP.Sheets(SheetNew)
        Set P = DB_MATCH.Sheets(Header).Range("HDR_" & SheetNew)
        NewEOL = EOL(SheetNew, DB_TMP) + 1
        For i = 1 To P.Columns.Count
            Width = Split(P.Cells(PTRN_WIDTH, i), "/")
            '------------------------
            X = X_Parse(SheetDB, SheetNew, P.Cells(PTRN_COLS, i), DB_Line, NewEOL, ExtPar)
            '--------------------------
'            sX = P.Cells(PTRN_COLS, i)
'            If sX <> "" Then
'                If InStr(sX, EXT_PAR) > 0 Then
'                    sXarr = Split(sX, "/")
'                    If UBound(sXarr) = 0 Then
'                        X = ExtPar
'                    Else
'                        If Not IsNumeric(sXarr(1)) Then
'ErrExtPar:                  ErrMsg FATAL_ERR, "Bad ExtPar: '" & sX & "'"
'                            End
'                        End If
'                        If UBound(ExtPar) < CLng(sXarr(1)) Then GoTo ErrExtPar
'                        X = ExtPar(sXarr(1))
'                    End If
'                ElseIf Left(sX, 1) = "#" Then
'                    sX = Mid(sX, 2)
'                    X = Workbooks(Rnew.RepFile).Sheets(Rnew.SheetN).Cells(Rnew.EOL, CLng(sX))
'                Else
'                    X = Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN).Cells(DB_Line, CLng(sX))
'                End If
                
                Y = Adapter(P.Cells(PTRN_ADAPT, i), X, P.Cells(PTRN_FETCH, i), IsErr, , DB_Line)
                
                If IsErr Then
                    MS "WrNewSheet: Ошибка при записи в лист '" & SheetNew & "'" _
                        & vbCrLf & " из листа '" & SheetDB & "' строки=" & DB_Line _
                        & vbCrLf & " в колонке=" & i & " =" & P.Cells(PTRN_COLS, i) & " X=" & X _
                        & vbCrLf & " Адаптер=" & P.Cells(PTRN_ADAPT, i) _
                        & " Fetch =" & P.Cells(PTRN_FETCH, i)
                    .Rows(NewEOL).Delete
                    Exit For
                Else
'                    .Cells(Rnew.EOL, i) = y
                    '-- записываем в SheetNew значение Y с установкой формата вывода
                    fmtCell DB_TMP, SheetNew, Width, Y, NewEOL, i
                End If
'!            Else
'                .Cells(Rnew.EOL, i) = P.Cells(2, i) '!!!!!!!!!!!!!???????????!!!!!!!!!!!!
                '-- iX пустой - записываем в SheetNew значение из Шаблона в указанном формате
'!                fmtCell DB_TMP, SheetNew, Width, P.Cells(2, i), Rnew.EOL, i
'!            End If
        Next i
    End With
    If Not IsErr Then
'!        RepTOC = Rnew
        WrTOC SheetNew
    End If
End Sub
Sub testWP_Adapt()
    WP_Adapt "HDR_WP", 2
    Stop
End Sub
Sub WP_Adapt(ByVal F As String, ByVal iLine As Long)
'
' - WP_Adapt(F, iLine)    - выводит на экран WP по Платежу в iLine, подготавливая диалог,
'                           где F - прототип выводимой формы документа (WP_Prototype)
'
'   21.10.12
'   23.10.12 - X_Parse вынесен в отдельную подпрограмму
'    2.11.12 - вызов NewOpp если Select не нашел ни одного Проекта
'    9.11.12 - работа с Named Range WP
'   11.11.12 - введен глобальный флаг для отладки TraceWidth
'    7.12.12 - введены форматы вывода "Dbl", "Txt", "Date" в строке "width"
'   19.01.13 - вызвана setColWidth
'   03.09.13 - смена имени с xAdapt на WP_Adapt
'   12.09.13 - OppSelect interface changed

    Const WP_PROTOTYPE = "WP_Prototype"

    Dim R As TOCmatch                           ' обрабатываемый Документ
    Dim iRow As Integer, iCol As Integer        ' строка и колонка Шаблона F
    Dim PtrnType As String                      ' поле Тип Шаблона
    Dim PutToRow As Long, putToCol As Long
    Dim X As String                             ' параметр Адаптера
    Dim Rqst As String                          ' строка - обращение к Адаптеру
    Dim F_rqst As String                        '
    Dim Y As String
    Dim IsErr As Boolean
    Dim iSelect As Long     '''', WP_Row As Long
    Dim i As Long
    Dim LocalTOC As TOCmatch, WP_Prototype_Lines As Long
    Dim nOpp() As Long, QtyOpp As Long          ' массив номеров строк выбранных Проектов
            
'---- Создаем заново лист WP
    
    LocalTOC = GetRep(WP)
''    Set DB_TMP = FileOpen(F_TMP)
    With Workbooks(LocalTOC.RepFile)
        Application.DisplayAlerts = False
        On Error Resume Next
        .Sheets(WP).Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
        DB_MATCH.Sheets(WP_PROTOTYPE).Copy Before:=.Sheets(1)
        .Sheets(1).Name = WP
    End With
'===== Заполняем лист WP
    With DB_TMP.Sheets(WP)
        .Tab.Color = rgbCoral
            '-- стираем WP_Prototype и заполняем новый шаблон,
            '   .. чтобы изменить его, сохранив Select-класс
        For i = 1 To EOL(WP, DB_TMP)
            .Rows(1).Delete
        Next i
        
        Dim FF As Range:  Set FF = DB_MATCH.Sheets(WP_PROTOTYPE).Range(F)
        Dim Width() As String
        FF.Copy .Cells(1, 1)
        .Cells(1, 5) = "'" & DirDBs & F_MATCH & "'!WP_Adapt_Continue"
'---- задаем ширину и заголовки вставленных колонок
        For i = 1 To FF.Columns.Count
'            If Not TraceWidth Then .Columns(i).ColumnWidth = FF.Cells(3, i)
            If Not TraceWidth Then setColWidth DB_TMP.Name, WP, i, FF.Cells(3, i)
        Next i
        
        .Cells(WP_CONTEXT_LINE, WP_CONTEXT_COL) = iLine
        WP_Prototype_Lines = EOL(WP, DB_TMP)
        For iRow = 1 To WP_Prototype_Lines Step PTRN_LNS
            iSelect = 0: QtyOpp = 0
            Do
                PtrnType = .Cells(iRow, 2)
                If PtrnType = PTRN_SELECT Then
                    nOpp = OppSelect(.Cells(8, 4))
                    QtyOpp = nOpp(0)
                    If QtyOpp = 0 Then GoTo StripEnd
                    iLine = nOpp(iSelect)   ' вывод Проектов по массиву индексов
                End If
                If .Cells(iRow, 1) <> "" Then
                    R = GetRep(.Cells(iRow, 1))
                    Workbooks(R.RepFile).Sheets(R.SheetN).Activate
                End If
                For iCol = 5 To .UsedRange.Columns.Count
                    X = X_ParseWP(iRow, iCol, PutToRow, putToCol, iLine)
                    Rqst = .Cells(iRow - 1 + PTRN_ADAPT, iCol)
                    F_rqst = .Cells(iRow - 1 + PTRN_FETCH, iCol)
                    
                    Y = Adapter(Rqst, X, F_rqst, IsErr, R.EOL, iRow, iCol)
                    
                    X = .Cells(iRow + PTRN_COLS - 1, iCol)
                    If X = "-1" Then Exit For
                    If Not IsErr And X <> "" Then
                        Width = Split(.Cells(iRow + PTRN_WIDTH - 1, iCol), "/")
                        fmtCell DB_TMP, WP, Width, Y, PutToRow + iSelect, putToCol
                    End If
                Next iCol
                iSelect = iSelect + 1
            Loop While iSelect < QtyOpp
                
StripEnd:   .Rows(iRow - 1 + PTRN_COLS).Hidden = True
            .Rows(iRow - 1 + PTRN_ADAPT).Hidden = True
            .Rows(iRow - 1 + PTRN_WIDTH).Hidden = True
            .Rows(iRow - 1 + PTRN_FETCH).Hidden = True
            If PtrnType = PTRN_SELECT And QtyOpp = 0 Then
                .Rows(iRow + 1).Hidden = True
                .Cells(iRow + 1 + PTRN_LNS, 11) = "В Salesforce нет подходящих Проектов. " _
                    & "Поэтому нажмите одну из кнопок [NewOpp], [->] или [STOP]"
                .Rows(iRow + 1 + PTRN_LNS).Interior.Color = rgbRed
            End If
        Next iRow
    End With
    DB_TMP.Sheets(WP).Activate
    
    
'''    If iSelect = 2 And Y = "-1" Then
'''       WP_Adapt_Continue "NewOpp", 1
'''    End If

'''''''''''''''''''''''''''''''''''
    End '''  остановка VBA ''''''''
'''''''''''''''''''''''''''''''''''
End Sub
Sub WP_Adapt_Continue(Button As String, iRow As Long)
'
' * WP_Adapt_Continue(Button) - продолжение работы WP_Adapt после нажатия кнопки Button
'                             Сюда передается управления из WP_Select_Button.
' 8/10/12
' 20.10.12 - обработка кнопок "Занести"
' 10.11.12 - bug fix - рекурсивный вызов WP с неправильным Namer Range

    Dim Proc As String, Step As String, iStep As Long
    Dim iPayment As Long, OppId As String, IsErr As Boolean
    Dim AccId As String, DefName As String
    Dim Respond As String, PaymentGoodType As String, PaymentContract As String
'---- извлекаем контектст из листа WP, то есть строки Платежа, Проекта -----
    With ActiveSheet
        iPayment = .Cells(WP_CONTEXT_LINE, WP_CONTEXT_COL)
        OppId = .Cells(iRow, 6)
        AccId = .Cells(WP_CONTEXT_LINE, 5)
        PaymentGoodType = .Cells(WP_CONTEXT_LINE, 11)
        PaymentContract = ContrCod(.Cells(WP_CONTEXT_LINE, 21), .Cells(WP_CONTEXT_LINE, 22))
    End With
    
    If DB_TMP Is Nothing Then Set DB_TMP = FileOpen(F_TMP)
    If DB_1C Is Nothing Then Set DB_1C = FileOpen(F_1C)
    If DB_MATCH Is Nothing Then Set DB_MATCH = FileOpen(F_MATCH)
    
    With DB_MATCH.Sheets(Process)
        .Activate
        Proc = .Cells(1, PROCESS_NAME_COL)
        Step = .Cells(1, STEP_NAME_COL)
        iStep = ToStep(Proc, Step)
''        .Cells(iStep, PROC_PAR2_COL) = iPayment + 1
    End With
    
    Select Case Button
    Case "STOP":
        StepOut Step, iStep
        ProcStart Proc
        End
    Case "->":
    Case "NewOpp":
        Const FETCH_SFACC = "SFacc/" & SFACC_IDACC_COL & ":" & SFACC_ACCNAME_COL
        DefName = FetchDoc(FETCH_SFACC, AccId, IsErr) & "-" & PaymentGoodType _
            & " Договор " & PaymentContract
        Respond = MsgBox(DefName, vbYesNo, "Имя нового Проекта")
        If Respond = vbYes Then
            WrNewSheet NEW_OPP, PAY_SHEET, iPayment, DefName
        End If
        Paid1C iPayment + 1
    Case "NewAcc":
    ' пока не написано
'-------- Обработка кликов на кнопках строк Select
    Case "Занести":
        WrNewSheet NEW_PAYMENT, PAY_SHEET, iPayment, OppId
    Case "Связать  ->"
        MS "->"
        Stop
        WrNewSheet DOG_UPDATE, PAY_SHEET, iPayment
    End Select
    
NextWP:         ProcStart Proc

End Sub
Sub Adapt(Optional FromDoc As String = "", Optional ToDoc As String = "")
'
' S Adapt([FromDoc],[ToDoc]) - запускает Адаптеры из Шаблона, осуществляя проход
'                              по Документу. Документ - ActiveSheet - первый в
'                              списке Документов Шага.
'                   Опциональные FromDoc и ToDoc - входной и выходной Документы.
'                   По умолчанию чтение и запись происходят в основной Документ.
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
'  3.1.13 - введено профилирование
'  6.1.13 - Optional FromDoc и ToDoc - по умолчанию ActiveSheet
' 10.1.13 - наличие ToDoc - признак записи в новый лист
' 24.1.13 - вызов fmyCell для записи Y вместе с форматом вывода
'  2.5.13 - в строке Шаблона Column теперь допустима ссылка на выходное поле #n
' 24.8.13 - Основной Документ- активный, а его Шаблон - атрибут в TOC
' 31.8.13 - для ускорения работы Шаблона выходим, если Процесса нет в ForProcess
'  1.8.13 - если ThisProc - не обрабатываем Шаблон слева от колонки <ThisProcess>

    StepIn
    
    Dim FF As Range     '= Форма F
    Dim R As TOCmatch
    Dim Rqst As String, F_rqst As String, IsErr As Boolean
    Dim X As String, Y As String
    Dim i As Long, Col As Long, iX As Long, iTo As Long, sX As String
    
    Dim R_From As TOCmatch, R_To As TOCmatch
    Dim Width() As String
    Dim Proc As String  ' поле в Шаблоне для проверки имени Процесса
    Dim ThisProcCol As Long, iProc As Long, sProc() As String, IsThisProc As Boolean

    ' профилирование
    
    Dim tot1 As Single, beg1 As Single: tot1 = 0
    Dim tot2(40) As Single, beg2(40) As Single
    Dim profileStr As String
    
    R = GetRep(ActiveSheet.Name)
    
    Set FF = DB_MATCH.Sheets(Header).Range( _
        DB_MATCH.Sheets(TOC).Cells(R.iTOC, TOC_FORMNAME))
    If FromDoc = "" Then
        R_From = R
    Else
        R_From = GetRep(FromDoc)
    End If
    If ToDoc = "" Then
        R_To = R
        iTo = 2
    Else
        R_To = GetRep(ToDoc)
        NewSheet ToDoc
        iTo = R_To.EOL
    End If
    Proc = "": ThisProcCol = 0: IsThisProc = False  '=инициируем для выяснения Процесса по Шаблону
    
    beg1 = Timer()                  ' профилирование
    For i = 2 To R_From.EOL
        Progress i / R_From.EOL
        iTo = iTo + 1
        If ToDoc = "" Then iTo = i
        For Col = 1 To FF.Columns.Count
            beg2(Col) = Timer()       ' профилирование
                            ' если ThisProc - пропускаем левую часть Шаблона
            If IsThisProc And Col < ThisProcCol Then Col = ThisProcCol
         '--- подготовка X - параметра Адаптера
            sX = FF(PTRN_COLS, Col)
            If sX = "" Then GoTo NextCol
            If IsNumeric(sX) Then
                iX = sX
                If iX < 0 Then Exit For
                X = Workbooks(R_From.RepFile).Sheets(R_From.SheetN).Cells(i, iX)
            ElseIf Left(sX, 1) = "#" Then
                sX = Mid(sX, 2)
                If Not IsNumeric(sX) Or CLng(sX) < 0 Then
                    ErrMsg FATAL_ERR, "Bad Column in Adapter ='" & sX & "'"
                    End
                End If
            ElseIf FF(PTRN_HDR, Col) = FOR_PROCESS Then
              '-- проверка контекста Процесса - заполнение IsThisProc
                If Proc = "" Then               'только первый раз
                    Proc = FF(PTRN_VALUE, Col)
                    sProc = Split(Proc, ",")
                    For iProc = LBound(sProc) To UBound(sProc)
                        If sX = PublicProcName Then
                            IsThisProc = True
                            ThisProcCol = Col + 1
                            If FF(PTRN_COLS, Col) = PublicProcName Then GoTo NextCol
                            ErrMsg FATAL_ERR, "Ошибка в Шаблоне '" _
                                & DB_MATCH.Sheets(TOC).Cells(R.iTOC, TOC_FORMNAME) & "'" _
                                & " в Col=" & Col & vbCrLf & "Ожидалось имя Процесса '" _
                                & PublicProcName & "', а не '" & FF(PTRN_COLS, Col) & "'"
                            End
                        End If
                    Next iProc
                End If
                If Not IsThisProc Then GoTo NextRow
'''            Else
'''                iX = sX
'''                X = Workbooks(R_To.RepFile).Sheets(R_To.SheetN).Cells(i, iX)
            End If
            
          '--- вызов Адаптера
            Rqst = FF.Cells(PTRN_ADAPT, Col)
            F_rqst = FF.Cells(PTRN_FETCH, Col)

            Y = Adapter(Rqst, X, F_rqst, IsErr, R_From.EOL, i, Col)

            If IsErr Then
                iTo = iTo - 1
                Exit For
            End If
          '--- записываем в SheetNew значение Y с установкой формата вывода
            Width = Split(FF.Cells(PTRN_WIDTH, Col), "/")
            fmtCell Workbooks(R_To.RepFile), R_To.SheetN, Width, Y, iTo, Col
            
NextCol:    tot2(Col) = tot2(Col) + (Timer() - beg2(Col))   ' профилирование
        Next Col
NextRow:
    Next i
'если ошибка в Адаптере NewSheet последней строки, тогда IsErr остается=True - стираем эту строку
    If IsErr And ToDoc <> "" Then Workbooks(R_To.RepFile).Sheets(R_To.SheetN).Rows(iTo).Delete

    ' профилирование
    tot1 = tot1 + (Timer() - beg1)
    profileStr = ""
    For Col = 1 To FF.Columns.Count
        profileStr = profileStr & " " & Format(tot2(Col), "###0.00")
    Next Col
    LogWr "adapt profile: total = " & Format(tot1, "###0.00") _
        & vbCrLf & "By steps = " & profileStr
End Sub
Function Adapter(ByVal Request As String, ByVal X As String, ByVal F_rqst As String, _
    ByRef IsErr As Boolean, Optional ByVal EOL_Doc As Long, _
    Optional ByVal iRow As Long, Optional ByVal iCol As Long) As String
'
' - Adater(Request, X, F_rqst, IsErr) - обрабатывает X в Адаптере "Request"
'    с внешними данными в Документе F_rqst. IsErr=True - ошибка в Адаптере
' 4.9.12
' 6.9.12 - bug fix
'25.9.12 - Dec(CurRate)
' 3.10.12 - Адаптер GetCol с синтаксисом ' GetCol/1C.xlsx,Платежи,5/SF:2:11
'12.10.12 - Адаптер GoodType(X)
'14.10.12 - Адаптер OppFilter для Шаблона типа Select
'18.10.12 - в OppFilter обработка EOL
'23.10.12 - CopyToVal и CopyFrVal
'25.10.12 - очистка переменных, оставшихся от прежних редакций
'18.11.12 - изменение кнопок "Связать"/"Занести"
'19.11.12 - добавление Адаптеров в колонки Шаблона справа - в т.ч. InvN
'29.12.12 - Compressor(X)
' 4.1.13 - Адаптер OppName для Платежей; обработка параметров Array
' 5.1.13 - Адаптер <>"" и <>1; выделение Адаптеров WP в отдельный модуль
' 7.1.13 - Изменения в GoodType - работа с флагами Лицензий, Подписки, Работ
'10.1.13 - Адаптер "Литерал; исправления TypeSFopp
'23.1.13 - новые Адаптеры IsBalky и BalkyOppId
' 7.4.13 - Адаптеры для БТО: BTO_Date, BTO_Order, BTO_Ord
'23.4.13 - Адаптер GetInv1C
'11.8.13 - Адаптеры "SN+" и "=Registered"
' 1.9.13 - все параметры ByVal, кроме возвращаемого ByRef IsErr
' 9.9.13 - добавлены Адаптеры TypOpp, LineOpp, KindOpp
'14.9.13 - добавлен Адаптер AltFetch: Y=X или его альтернатива из Fetch, если есть

    Dim FF() As String, tmp() As String, InitX As String
    Dim i As Long, Par() As String, Z(10) As String
    Dim WP_Row As Long  ' строка для записи результат Адаптеров, использется в Select
    Dim LocalTOC As TOCmatch
    
    IsErr = False
    X = Compressor(X)
    InitX = X
    
'--- разбор строки Адаптера вида <Имя>/C1,C2,C3...
    Dim AdapterName As String
    AdapterName = ""
    If Request <> "" Then
        tmp = Split(Request, "/")
        AdapterName = tmp(0)
        If InStr(Request, "/") <> 0 Then
            Par = Split(tmp(1), ",")
'!!'            Call ArrayZ(Z, PAY_SHEET, iRow, Par)
        End If
    End If

'======== препроцессинг Адаптера для подварительной обработки X перед Fetch =========
    Select Case AdapterName
    Case "MainContract":
        X = Trim(Replace(X, "Договор", ""))
    Case "<>0", "SN+":
        If X = "0" Then X = ""
    Case "<>"""""
        If X = "" Then
            IsErr = True
            Exit Function
        End If
    Case "=Registered"
        If X <> "Registered" Then
            IsErr = True
            Exit Function
        End If
    Case "ContrK":
        Const PAY_REF = 8
        Dim MainDog As String, iPay As Long
        iPay = DB_TMP.Sheets(WP).Cells(PAY_REF, 4)
        MainDog = DB_1C.Sheets(PAY_SHEET).Cells(iPay, CLng(Par(0)))
        X = ContrCod(X, MainDog)
    Case "AltFetch":    ' Y=X или его альтернатива из Fetch, если есть
        Dim Str As String
        Str = FetchDoc(F_rqst, X, IsErr)
        If Str <> "" Then X = Str
        GoTo AdapterSelect
    End Select
    
'--- FETCH разбор строки параметров из Документов вида <Doc1>/C1:C2,<Doc2>/C1:C2,...
    If F_rqst <> "" And X <> "" Then
        FF = Split(F_rqst, ",")
        For i = LBound(FF) To UBound(FF)
            X = FetchDoc(FF(i), X, IsErr)
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' сейчас используется только один указатель на извлекаемую из Doc величину.
' В дальнейшем надо использовать массив x(1 to 5) и обращаться к Fetch несколько раз
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Next i
    End If

'******* выполнение Адаптера с параметрами Par ******
''    Set DB_TMP = FileOpen(F_TMP)
''    With DB_TMP.Sheets(WP)
AdapterSelect:

    Adapter = ""
'''''    If SheetWP = WP Then
'''''        Adapter = AdpterWP()
'''''        Exit Function
'''''    End If
    
    Select Case AdapterName
    Case "", "MainContract", "<>""""", "=Registered", "AltFetch": Adapter = X
    Case "Мы", "Продавцы", "Продавец_в_SF", "Vendor":
        On Error GoTo AdapterFailure
        Adapter = WorksheetFunction.VLookup(X, DB_MATCH.Sheets("We").Range(AdapterName), Par(0), False)
        On Error GoTo 0
    Case "Dec": Adapter = Dec(X)
    Case "GetCol":
        If X <> "" Then           ' GetCol/1C.xlsx,Платежи,5 [/SF/2:11]
            Adapter = Workbooks(Par(0)).Sheets(Par(1)).Cells(CLng(X), CLng(Par(2)))
            If UBound(tmp) > 1 Then
                Adapter = FetchDoc(tmp(2) & "/" & tmp(3), Adapter, IsErr)
            End If
        End If
    Case "GoodType": Adapter = GoodType(X)
    Case "GoodJob":
        Call ArrayZ(Z, PAY_SHEET, iRow, Par)
        If GoodJob(Z(1), X, CLng(Par(0))) Then Adapter = "1"
    Case "CurISO":  Adapter = CurISO(X)
    Case "CurRate": Adapter = Dec(CurRate(X))
    Case "Дата":    If X <> "" Then Adapter = DDMMYYYY(X)
    Case "ContrK":  Adapter = X 'преобразование в вид ContrCod в препроцессинге
    Case "SeekInv": Adapter = SeekInv(X)
    Case "InvN":
        tmp = Split(X, " ")
        If UBound(tmp) > 0 Then Adapter = tmp(0)
    Case "SeekPayN":
        Dim Inv As String, Client As String
        Inv = ActiveSheet.Cells(iRow, CLng(Par(0)))
        Client = ActiveSheet.Cells(iRow, CLng(Par(1)))
        If Inv <> "" And IsDate(X) Then Adapter = SeekPayN(Inv, Client, X)
        If Adapter = "0" Then Adapter = ""
    Case "<>1":
        Call ArrayZ(Z, PAY_SHEET, iRow, Par)
        If Z(0) = "1" Then
            IsErr = True
        Else
            Adapter = X
        End If
    Case "SN+":
        Adapter = X & "+"
    Case "OppName":
        If X <> "" Then
            Adapter = X
        Else
            Call ArrayZ(Z, PAY_SHEET, iRow, Par)
            Adapter = Z(2) & "-" & Z(3) & " " & ContrCod(Z(4), Z(5))
        End If
    Case "BTO_Date":
        Adapter = Mid(X, 2, WorksheetFunction.FindB(" ", X) - 2)
    Case "BTO_Order":
        Dim ChBeg As Long, ChEnd As Long
        ChBeg = WorksheetFunction.FindB("по счету", X) + 9
        ChEnd = WorksheetFunction.FindB(" ", X, ChBeg)
        Adapter = Mid(X, ChBeg, ChEnd - ChBeg)
    Case "TypeSFopp":
        Call ArrayZ(Z, PAY_SHEET, iRow, Par)
        If Z(0) = "1" Then Adapter = "Лицензии"
        If Z(1) = "1" Then Adapter = "Подписки"
        If Z(3) = "1" Then Adapter = "Работы"
        If X = "Расходники" Then Adapter = X
        If X = "Оборудование" Then Adapter = "Железо"
    Case "LineOpp":
        Adapter = "Программное обеспечение (ПО)"
        Select Case X
        Case "Расходники": Adapter = "Расходные материалы и ЗИП"
        Case "Работы":     Adapter = "Услуги"
        Case "Печать":     Adapter = "Печать"
        Case "Оборудование": Adapter = "железо"
        End Select
    Case "KindOpp":
        Dim GoodTp As String
        GoodTp = GoodType(X)
        On Error Resume Next
        Adapter = WorksheetFunction.VLookup(GoodTp, DB_MATCH.Sheets(We).Range("OppTypeRng"), 4, False)
        On Error GoTo 0
    Case "Max":
        Call ArrayZ(Z, PAY_SHEET, iRow, Par)
        Adapter = X
        For i = LBound(Z) To UBound(Z)
            Adapter = WorksheetFunction.Max(CLng(Adapter), CLng(Z(i)))
        Next i
    Case "EmptyBuddy":
        On Error GoTo AdapterFailure
        Adapter = WorksheetFunction.VLookup(X, DB_MATCH.Sheets("We").Range("Продавцы"), 3, False)
        On Error GoTo 0
        If InStr(Adapter, X) = 0 Then Adapter = X
    Case "ForceTxt":
        Adapter = "'" & X
    Case "DogVal":                                      '=Max(Платежа, Счета, Суммы Договора)
        Dim Vpaid As Long, Vinv As Long, Vdog As Long   ' величины Платежа, Счета и Договора
        Dim sDog As String, DogCur As String            ' имя Договора и его валюта
        LocalTOC = GetRep(PAY_SHEET)
        With Workbooks(LocalTOC.RepFile).Sheets(LocalTOC.SheetN)
            Vpaid = .Cells(iRow, PAYRUB_COL)
            Vinv = .Cells(iRow, PAYINVVAL_COL)
            sDog = .Cells(iRow, PAYDOGVAL_COL)
            DogCur = .Cells(iRow, PAYDOGCUR_COL)
        End With
        If sDog <> "" Then
            If Not IsNumeric(sDog) Then
                ErrMsg FATAL_ERR, "Не числовое значение в сумме Договора по " & X
                Stop
                End
            End If
            Vdog = CDbl(sDog) * CurRate(DogCur)
        End If
        Adapter = Dec(Application.Max(Vpaid, Vinv, Vdog))
    Case "TypOpp":  '!!' сменить название
    ' -- распознавание типа Проекта по типу и спецификации Товара
'        Dim Good As String
'        LocalTOC = GetRep(PAY_SHEET)
'        With Workbooks(LocalTOC.RepFile).Sheets(LocalTOC.SheetN)
'            Good = .Cells(iRow, PAYGOOD_COL)
'        End With
        Adapter = TypOpp(X)
    Case "CopyToVal", "CopyFrVal", "OppType", "OppFilter", _
            "SetOppButton", "NewOppNameFromWP":
        Adapter = AdapterWP(AdapterName, X, Par)
    Case "IsBalky":
        Call ArrayZ(Z, PAY_SHEET, iRow, Par)
        If Z(0) = "" Or Z(1) = "1" Or Z(2) <> "Расходники" Then
            IsErr = True
        Else
            Adapter = X
        End If
    Case "BalkyOppN":  'SFopp/4:2 с проверкой колонки OpportunityActivityKind = "Расходники"
        Dim BalkyExists As Boolean: BalkyExists = False
        Dim Rdoc As TOCmatch, N As Long, FromN As Long
        Rdoc = GetRep(SFopp)
        FromN = 1
        With Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN)
            Do
                N = CSmatchSht(X, SFOPP_ACC1C_COL, Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN), FromN)
                If N = 0 Then
                    If BalkyExists Then Exit Function
                    GoTo AdapterFailure
                End If
                Dim OppCloseDate As Date, DaysToClos As Long
                OppCloseDate = .Cells(N, SFOPP_CLOSEDATE_COL)
                DaysToClos = OppCloseDate - Now
                If .Cells(N, SFOPP_LINE_COL) = OPP_BALKY And DaysToClos > 365 Then
                    If BalkyExists Then
                    End If
                    BalkyExists = True
                    Adapter = .Cells(N, SFOPP_OPPN_COL)
                End If
                FromN = N + 1
            Loop
        End With
    Case Else
        If Left(AdapterName, 1) = """" Then
            Adapter = Mid(AdapterName, 2)
        Else
            ErrMsg FATAL_ERR, "Adapter> Не существует " & AdapterName
        End If
    End Select
''    End With
    Exit Function
    
AdapterFailure:
    ErrMsg WARNING, "Адаптер " & AdapterName & "(" & X & ") не получил данных"
SkipLine:
    IsErr = True
End Function
Function AdapterWP(AdapterName, X, Par) As String
'
' - AdapterWP() - обработка Адаптеров для Шаблонов WP
' 5.1.2013
' 4.9.13 - дописываем

    Dim i As Long, Z(10) As String

    ''??'' WP_PAYMENT_LINE =?= iRow
    
    Set DB_TMP = FileOpen(F_TMP)
    With DB_TMP.Sheets(WP)
    
        Select Case AdapterName
'''        Case "CopyToVal":
'''            Set DB_TMP = FileOpen(F_TMP)
'''            With DB_TMP.Sheets(WP)
'''                WP_Row = iRow + .Cells(iRow + 3, 3) + PTRN_LNS - 1
'''               .Cells(iRow - 1 + PTRN_VALUE, iCol).Copy .Cells(WP_Row, iCol)
'''            End With
'''        Case "CopyFrVal":
'''            Set DB_TMP = FileOpen(F_TMP)
'''            With DB_TMP.Sheets(WP)
'''                WP_Row = iRow + .Cells(iRow + 3, 3) + PTRN_LNS - 1
'''                .Cells(WP_Row, iCol).Copy .Cells(iRow - 1 + PTRN_VALUE, iCol)
'''            End With
'''        Case "OppType":             ' инициализация типа Проекта
'''        ''''        Call ArrayZ(Z, PAY_SHEET, iRow, Par)
'''            If X = "Оборудование" Then X = "Железо"
'''            If X = "Autodesk" Then
'''        ''  !          If IsSubscription(.Cells(WP_PAYMENT_LINE, CLng(Par(0))), X) Then
'''                If IsSubscription(Z(0), X) Then
'''                    X = "Подписка"
'''                Else
'''                    X = "Лицензии"
'''                End If              '!' не рассматриваются Работы!!
'''            End If
'''            AdapterWP = X
        Case "OppFilter":
            Const SEL_REF = 20
        ' проверить есть ли Проект связанный с Договором
            Dim IdSFopp As String
            IdSFopp = .Cells(SEL_REF, 3)
            If IdSFopp = "" Then
                Dim b As Long, a(0 To 6) As Long
                b = .Cells(SEL_REF + 2, 4)
                For i = 0 To UBound(a)
'                    a(i) = CLng(Par(i))
                Next i
                AdapterWP = "-1"  ' -1 - признак, что достигнут EOL, и Проект не найден
'''        !            For i = .Cells(SEL_REF, 4) + 1 To EOL_Doc
'''        !                If OppFilter(i, .Cells(b, A(0)), .Cells(b, A(1)), _
'''        !                        .Cells(b, A(2)), .Cells(b, A(3)), .Cells(b, A(4)), _
'''                            .Cells(b, A(5)), .Cells(b, A(6))) Then
'''                        Adapter = i
'''                        Exit For
'''                    End If
'''                Next i
            Else
        ' вывести один единственный Проект, когда Платеж с Договором, и он связан с Проектом
                Dim Rdoc As TOCmatch, Doc As String
''        !Doc = .Cells(iRow, 1)
                Rdoc = GetRep(Doc)
                AdapterWP = CSmatchSht(X, SFOPP_OPPID_COL, Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN))
'''        !            .Cells(iRow + PTRN_LNS, 11) = "Занести"
'''        !            .Cells(iRow + PTRN_LNS, 11).Interior.Color = rgbBlue
'''        !            If Adapter = .Cells(iRow + 1, 4) Then Adapter = "-1"
            End If
'''        Case "SetOppButton":
'''        !WP_Row = iRow + .Cells(iRow + 3, 3) + PTRN_LNS - 1          ' копирование кнопки "Связать"
'''        !        .Cells(iRow - 1 + PTRN_VALUE, iCol).Copy .Cells(WP_Row, iCol)
'''            If X = "" Then Adapter = "Занести"  ' Если в Платеже нет Договора - кнопка "Занести"
''''''        Case "NewOppNameFromWP":
''''''        ' -- формируем имя Проекта в виде Организация-ТипТовара Договор Дата
''''''            Dim Typ As String, Dogovor As String, Dat As String
''''''        ''!Typ = .Cells(WP_PAYMENT_LINE, CLng(Par(0)))
''''''        ''!Dogovor = .Cells(WP_PAYMENT_LINE, CLng(Par(1)))
''''''        ''!MainDog = .Cells(WP_PAYMENT_LINE, CLng(Par(2)))
''''''        ''        Dogovor = ContrCod(Dogovor, MainDog)
''''''        ''!Dat = .Cells(WP_PAYMENT_LINE, CLng(Par(3)))
''''''            Typ = Z(0): Dogovor = Z(1): MainDog = Z(2): Dat = Z(3)
''''''            Dogovor = ContrCod(Dogovor, MainDog)
''''''            AdapterWP = X & "-" & Typ & " " & Dogovor & " " & Dat
        End Select
    End With
End Function
Sub testArrayZ()
    Dim Z(3) As String, Par(3) As String
    Par(0) = "9"
    Par(1) = "31"
    Par(2) = "25"
    Par(3) = "26"
    Dim t0, t1, t2
    t0 = Now
    Call ArrayZ(Z, PAY_SHEET, 2886, Par)
    t1 = Now
    Call ArrayZ(Z, PAY_SHEET, 3082, Par)
    t2 = Now
    t2 = 1000 * (t2 - t1)
    t1 = 1000 * (t1 - t0)
End Sub
Sub ArrayZ(ByRef Z() As String, Doc As String, iRow As Long, ByRef Par() As String)
'
' - ArrayZ(Z(), Doc, iRow, Par())   - по массиву номеров колонок в Par возвращает в массив Z()
'                               значений из ячеек документа Doc в строке iRow в колонках Par()
' 4.1.13

    Dim R As TOCmatch, i As Long
    
    R = GetRep(Doc)
    With Workbooks(R.RepFile).Sheets(R.SheetN)
        For i = LBound(Par) To UBound(Par)
            Z(i) = Trim(.Cells(iRow, CLng(Par(i))))
        Next i
    End With
End Sub
Sub TestX_Parse()
'
' T TestX_Parse()
'
'   8.9.13

    Dim X

    X = X_Parse(X_rqst:="7", InDoc:=PAY_SHEET) ' должен вернуть Date из EOL Платежей
    X = X_Parse(X_rqst:="#1", OutDoc:=NEW_OPP) ' должен вернуть значение из колонки 1 EOL NewOpp
' должен вернуть значение из колонки 4 заголовка выходного листа NewOpp
    X = X_Parse(X_rqst:="#4", OutDoc:=NEW_OPP, iLine:=1)
    X_Parse X_rqst:=EXT_PAR ' должен дать FATAL_ERR, а он гад возвращает ""
    X_Parse X_rqst:=EXT_PAR, ExtPar:="проверочный текст"    ' возвращает строку из ExtPar
End Sub
Function X_Parse(Optional InDoc As String, Optional OutDoc As String = "", _
    Optional X_rqst As String, Optional iLine, Optional PutToRow, _
    Optional ExtPar As String) As String
'
' -  X_Parse([InDoc,OutDoc,X_rqst,iLine,PutToRow,ExtPar)  - разбор строки Х-параметра Адаптера
'            из поля X_rqst Шаблона. По префиксу # в X_rqst X извлекается из OutDoc
'   здесь
'           X_rqst          - разбираемая строка параметра Адаптера X
'         (iLine,iCol)      - адрес ячейки Шаблона WP для разбора, ссылка на номер колонки.
'                             Если указаны (iLine, iCol), то значение X_rqst извлекается из WP
'       [PutToRow,PutToCol] - адрес ячейки, куда поместить результат Адаптера в WP
'           [iLine]         - номер строки для Шаблона WP типа iLine
'
' в поле Шаблона возможна конструкция #6 или !6, а также "ExtPar"
'  * знак # означает, что адресуется не колонка в ActiveSheet, а колонка самого Шаблона
'  * знак @ - что Х надо извлечь из Шаблона -- не дописано
'  * знак ! - что Х надо извлечь из строки WProw ниже Шаблона
'  * "ExtPar" - возвращает X_Parse значение параметра ExtPar
'  ? будет дописана конструкция вида #3/Format, где после "/" указана доп.информация
'
' 22.10.12
' 25.10.12 - иcправления в связи с HashFlag=True
' 11.11.12 - добавлен синтаксис !<Col> для адресации WProw
' 15.11.12 - Optional параметры
'  8.09.13 - обработка любого Шаблона из Optional параметр X_rqst, не только из WP
'          - Optional параметр ExtPar: если в поле Шаблона "ExtPar", Х=этому параметру

    Dim InDocTOC As TOCmatch, OutDocTOC As TOCmatch
    Dim ErrStr As String, FF As Range
    Dim sX() As String
    Dim iX As Long, WP_Row As Long
    Dim RefType As String
    Dim iCol As Long            '= номер колонки
    
    X_Parse = ""
    If IsMissing(InDoc) Then InDoc = ActiveSheet.Name
    If IsMissing(X_rqst) Or X_rqst = "" Then GoTo Ex
    ErrStr = "X_Parse> "
    
    If IsNumeric(X_rqst) Then
        If Trim(InDoc) = "" Then GoTo ErrInDoc
        InDocTOC = GetRep(InDoc)
        If IsMissing(iLine) Then iLine = InDocTOC.EOL
        iCol = X_rqst
        GoTo GetX
    ElseIf X_rqst = EXT_PAR Then
        If IsMissing(ExtPar) Then GoTo ErrExtPar
        X_Parse = ExtPar
        GoTo Ex
    Else
        RefType = Left(X_rqst, 1)
        If RefType = "#" Then
            iCol = Mid(X_rqst, 2)
            If IsMissing(OutDoc) Then OutDoc = InDoc
            If IsMissing(PutToRow) Or Not IsNumeric(PutToRow) Then GoTo ErrPutToRow
            OutDocTOC = GetRep(OutDoc)
            With OutDocTOC
                X_Parse = Workbooks(.RepFile).Sheets(.SheetN).Cells(PutToRow, iCol)
            End With
        End If
        GoTo Ex
    End If
    
GetX:
    If Not IsNumeric(iLine) Or iLine <= 0 Or iLine > InDocTOC.EOL Then GoTo Err_iLine
    If Not IsNumeric(iCol) Or iCol <= 0 Then GoTo Err_iCol
    
    X_Parse = Workbooks(InDocTOC.RepFile).Sheets(InDocTOC.SheetN).Cells(iLine, iCol)

Ex: Exit Function
ErrInDoc:       ErrMsg FATAL_ERR, ErrStr & "Недоступен InDoc"
ErrOutDoc:      ErrMsg FATAL_ERR, ErrStr & "Недоступен OutDoc"
ErrExtPar:      ErrMsg FATAL_ERR, ErrStr & "Недоступен ExtPar"
Err_iLine:      ErrMsg FATAL_ERR, ErrStr & "странный iLine=" & iLine
Err_iCol:       ErrMsg FATAL_ERR, ErrStr & "странный iCol=" & iCol
ErrPutToRow:    ErrMsg FATAL_ERR, ErrStr & "странный PutToRow=" & PutToRow
End Function
Function X_ParseWP(ByVal iRow, ByVal iCol, PutToRow, putToCol, _
    Optional iLine, Optional ExtPar As String) As String
'
'
'
'''''                X = X_Parse(iRow, iCol, putToRow, putToCol, iLine)

    Dim WP_TOC As TOCmatch, InDoc As String
    Dim iX As Long, WP_Row As Long
    Dim PtrnType As String
    Dim X_rqst As String, sX() As String
    
    WP_TOC = GetRep(WP)
    With Workbooks(WP_TOC.RepFile).Sheets(WP_TOC.SheetN)
        InDoc = .Cells(iRow, 1)
        WP_Row = iRow - 1 + PTRN_VALUE
        PtrnType = .Cells(iRow, 2)
        If PtrnType = PTRN_SELECT Then WP_Row = iRow + PTRN_LNS + .Cells(iRow + 3, 3) - 1
        PutToRow = WP_Row: putToCol = iCol
        X_rqst = .Cells(iRow - 1 + PTRN_COLS, iCol)
    End With
    
GetX:
    X_ParseWP = X_Parse(InDoc:=InDoc, OutDoc:=WP, X_rqst:=X_rqst, iLine:=iLine, PutToRow:=WP_Row)

'    If X_rqst = "" Then GoTo Ex
'
'    sX = Split(X_rqst, "/")
'    RefType = Left(sX(0), 1)
''''    If OutputDoc = WP Then      'если OutputDoc не указан, вся дальнейшая работа с WP
''''    If RefType = "#" Or RefType = "!" Then sX(0) = Mid(sX(0), 2)
'    If RefType = "#" Or RefType = "!" Then sX(0) = Mid(sX(0), 2)
'    X_Parse = sX(0)
''        If IsMissing(Form) Then GoTo ErrInDocForm
''        Set FF = DB_MATCH.Sheets(Header).Range(Form)
''        InDoc = FF(DB_MATCH.Cells(1, 1))
''''
''''    iX = 0
''''    If UBound(sX) >= 0 Then iX = sX(0)
''''    If iX > 0 Then
''''        Select Case PtrnType
''''        Case "Кнопки", "Шаблон": GoTo GetFromWP
''''        Case "iLine":
''''            WP_Row = iLine
''''            GoTo GetFromActiveSheet
''''        Case PTRN_SELECT:
''''            WP_Row = .Cells(putToRow, 5)
''''            GoTo GetFromActiveSheet
''''         Case Else:
''''            ErrMsg FATAL_ERR, "WP_Adapt> Странный тип Шаблона " & PtrnType
''''        End Select
''''        End If
''''
''''GetFromWP:
''''        If iX > 0 Then X_Parse = .Cells(WP_Row, iX)
''''        GoTo Ex
''''    End With
''''
''''GetFromActiveSheet:
''''    If RefType = "!" Then
''''        WP_Row = putToRow
''''        GoTo GetFromWP
''''    ElseIf RefType = "#" Then
''''        WP_Row = iRow + PTRN_VALUE - 1
''''        GoTo GetFromWP
''''    End If
''''    If iX > 0 Then X_Parse = ActiveSheet.Cells(WP_Row, iX)
'    If OutputDoc = WP Then      'если OutputDoc не указан, вся дальнейшая работа с WP
'        GetRep WP
'        With DB_TMP.Sheets(WP)
'            WP_Row = iRow - 1 + PTRN_VALUE
'
'            PtrnType = .Cells(iRow, 2)
'            If PtrnType = PTRN_SELECT Then WP_Row = iRow + PTRN_LNS + .Cells(iRow + 3, 3) - 1
'
'            putToRow = WP_Row: putToCol = iCol
'
'            X_rqst = .Cells(iRow - 1 + PTRN_COLS, iCol)
'        End With
'    End If
'    X_ParseWP = XParse(OutputDoc:=WP, iRow:=iRow, iCol:=iCol)
    End Function
Function FetchDoc(F_rqst, X, IsErr, Optional ByRef FromN As Long = 1) As String
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
'                               C2 может иметь значенте №, тогда по С1 возвращается
'                                  номер строки, где нашли значение С1
' *             Вторая группа - параметры обработки ошибок Fetch - /W или /0
' *  /W             - WARNING в Log, оставлять IsErr=False, если извлечено ""
' *  /0             - "" вполне допустимо (например, область в адресе)
' *  /D             - Default - "" допустим, но IsErr=True для значения по умолчанию
' *             Вторая группа отсутствует - выводить Log и IsErr = True
'
' 5.9.12
' 14.9.12 - работает /D для второй группы - "по умолчанию"
' 4.11.12 - Fetch возвращает номер строки в случае <Doc>/C1:№
' 3.09.13 - Optional FromN позволяет вести поиск не с начала Документа

    FetchDoc = ""
    If F_rqst = "" Or X = "" Then GoTo ErrExit
        
    Dim tmp() As String, Cols() As String, S As String
    Dim Doc As String, C1 As Long, C2 As Long, Rng As Range, N As Long
            
    tmp = Split(F_rqst, "/")
    Doc = tmp(0)
    Cols = Split(tmp(1), ":")
    C1 = Cols(0)
    
    Dim Rdoc As TOCmatch, W As Workbook
    Rdoc = GetRep(Doc)
    
    If UBound(Cols) < 1 Then
'--- ситуация С1 - в группе один параметр - извлекаем значение по индексу
        Dim Indx As Long
        Indx = X
'!!!!!!!!!!!!!!!!!!!!!!!!!!!
' сейчас Indx=Х - это просто число, но в дальнейшем тут надо split
'!!!!!!!!!!!!!!!!!!!!!!!!!!!
        If Indx <= 0 Then
            ErrMsg WARNING, "FetchDoc: " & Doc & "(" & Indx & "," & C1 _
                & ") - неправильный номер строки"
            GoTo ErrExit
        End If
        S = Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN).Cells(Indx, C1)
    Else
'--- ситуация С1:C2 - в группе 2 параметра - извлекаем значение по Lookup или №
        If IsNumeric(Cols(1)) Then C2 = Cols(1)
        S = ""
        N = CSmatchSht(X, C1, Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN), FromN)
        FromN = N
        If N <> 0 Then
            If Cols(1) = "№" Then
                S = N
            Else
                S = Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN).Cells(N, C2)
            End If
        End If
    End If
'--- обработка группы 2 -- если S=""
    If S = "" Then
        If UBound(tmp) >= 2 Then
            If tmp(2) = "W" Then
                ErrMsg WARNING, "Адаптер> ссылка " & F_rqst _
                    & "(" & X & ") не работает, результат <пусто>"
            End If
            If tmp(2) <> "0" Then GoTo ErrExit
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
Sub testfmtCell()
'   тесты fmtCell()
' 17.12.12 - добавлен тест целого формата
    
    Dim fmt(0 To 1) As String
    Set DB_TMP = FileOpen(F_TMP)
    
    fmt(1) = "Dbl"
    fmtCell DB_TMP, "NewOpp", fmt, "3m3", 2, 2
    fmtCell DB_TMP, "NewOpp", fmt, 33333.3, 2, 2
    
    fmt(1) = "Txt"
    fmtCell DB_TMP, "NewOpp", fmt, "xxx", 2, 2
    fmt(1) = "@"
    fmtCell DB_TMP, "NewOpp", fmt, "yyy", 2, 2
    
    fmt(1) = "Date"
    fmtCell DB_TMP, "NewOpp", fmt, "1/2/2012", 2, 2
    fmt(1) = "# ##0.0000"
    fmtCell DB_TMP, "NewOpp", fmt, 5666, 2, 2
    fmt(1) = "# ##0.00"             ' Стандарт для России ~ Dbl
    fmtCell DB_TMP, "NewOpp", fmt, 5666, 2, 2
    fmt(1) = "0.00"
    fmtCell DB_TMP, "NewOpp", fmt, 5666, 2, 2
    fmt(1) = "0"
    fmtCell DB_TMP, "NewOpp", fmt, 22, 2, 2
    fmt(1) = "0%"
    fmtCell DB_TMP, "NewOpp", fmt, 5.666, 2, 2
    fmt(1) = "0.00%"
    fmtCell DB_TMP, "NewOpp", fmt, 5.666, 2, 2
    Stop
End Sub
Sub fmtCell(ByVal db As Workbook, ByVal list As String, fmt() As String, _
            ByVal Value, ByVal PutToRow As Long, ByVal putToCol As Long)
'
'   - fmtCell - обработка формата в строке width вида 0/Txt или 10/@
'
' db и list - имена Workbook и листа, куда помещаем ..
' value     - значение в виде ..
' fmt(1)    - формата ячейки [putToRow,putToCol]
'
'   * распознаются форматы Txt, Dbl, Date и любые другие, понимаемые Excel
'
'  7.12.12 - введены форматы вывода "Dbl", "Txt", "Date" в строке "width" в sub WP_Adapt
'  8.12.12 - введен прoизвольный формат в строке width
' 17.12.12 - добавлен тест целого формата в testfmtCell()
' 19.12.12 - изменен разделитель троек в Dbl в testfmtCell()
' 17.12.12 - добавлен тест целого формата
' 12.9.13 - увеличено количество триад для Dbl

    If UBound(fmt) > 0 Then
        If fmt(1) = "Dbl" Then
'                                Dim YY As Double
'                                YY = Y
'                                .Cells(PutToRow, PutToCol) = YY
            db.Sheets(list).Cells(PutToRow, putToCol).NumberFormat = "# ### ##0.00"
        ElseIf fmt(1) = "Date" Then
            db.Sheets(list).Cells(PutToRow, putToCol).NumberFormat = "[$-409]d-mmm-yyyy;@"
        ElseIf fmt(1) = "Txt" Then
            db.Sheets(list).Cells(PutToRow, putToCol).NumberFormat = "@"
        Else
            db.Sheets(list).Cells(PutToRow, putToCol).NumberFormat = fmt(1)
        End If
    End If
    db.Sheets(list).Cells(PutToRow, putToCol) = Value
End Sub
