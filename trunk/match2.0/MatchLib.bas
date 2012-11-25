Attribute VB_Name = "MatchLib"
'---------------------------------------------------------------------------
' Библиотека подпрограмм проекта "match 2.0"
'
' П.Л.Храпкин, А.Пасс 25.11.2012
'
' - GetRep(RepName)             - находит и проверяет штамп отчета RepName
' - FatalRep(SubName, RepName)  - сообщение о фатальной ошибке при запросе RepName
' - WrTOC()                     - записывает Publoc RepTOC в TOCmatch
' - CheckStamp(iTOC, [FromMoveToMatch]) - проверка Штампа по стоке в TOCmatch
' - FileOpen(RepFile)           - проверяет, открыт ли RepFile, если нет - открывает
' S InsMyCol(F[,FS])            - вставляем колонки в лист слева по шаблону F и пятку из FS
' - MS(Msg)                     - вывод сообщения на экран и в LogWr
' - ErrMsg(ErrMode, MSG)        - вывод сообщения об ощибке в Log и на экран
' - LogWr(msg)                  - запись сообщения msg в Log list
'(*)LogReset()                  - сброс и очистка Log листа
' - ActiveFilterReset(SheetN)   - сброс и активизация автофильтра листа SheetN
' - SheetsCtrlH(SheetN, FromStr, ToStr) - замена текста FromStr на ToStr
'                                 в листе SheetN
' - PerCent(Row, Col)           - форматирование ячейки (Row,Col) с процентами
' - CurCode(Row, Col, CurCol)   - формат ячейки (Row,Col) по коду валюты в CurCol
' - CurRate(Cur)                - возвращает курс валюты к рублю по коду Cur для We
' - CurISO(Cur1C)               - возвращает код валюты в стандарте ISO
' - DDMMYYYY(d)                 - преобразование даты d в текстовый формат DDMMYYYY
' - Dec(a)                      - формат числа а в виде текста с десятичной точкой
' - EOL(SheetN)                 - возвращает номер последней строки листа SheetN
' - RowDel(RowStr)              - удаляет строки активного листа в соответствии с RowStr
' - CSmatch(Val,Col)            - Case Sensitive match - возвращает номер строки с Val
'                                 в колонке Col. Если Val не найден- возвращает 0.
'                                 Лист для поиска Val должен быть Selected.
' - CSmatchSht(Val,Col,Sht)     - Case Sensitive match возвращает номер строки с Val в
'                                 колонке Col листа Sht. Если Val не найден- возвращает 0.
' - ClearSheet(SheetN, HDR_Range) - очистка листа SheetN и запись в него шапки
' - SheetSort(SheetN, Col)      - сортировка листа SheetN по колонке Col
' - SheetDedup(SheetN, Col)     - cортировка и дедупликация SheetN по колонке Col
' - SheetDedup2(SheetN, ColSort,ColAcc) - сортировка и слияние листа SheetN
'                                 по колонкам ColSort, ColAcc
' - DateCol(SheetN, Col)        - преобразование колонки Col из текста в Дату
' - DateSort(SheetN, Col)       - преобразование колонки Col из текстового формата в Date
'                                 и сортировка по этой колонке от старых к новым датам
' - HideLns(FrR, ToR, Col, Criteria) - скрывает строки от FrR до ToR,
'                                 если Col соотв.Criteria (для WP)
' - Progress(Pct)               - вывод Progress Bar'а - процентов выполнения кода
' - StopSub()                   - аварийное завершение процесса - вызывается из формы
' - RemDelimeters(S)            - замена спец символов в строке S на пробелы
' - Compressor(S)               - удаление кратных пробелов из строки S
' Т testFindInLst()             - отладка FindInLst(W,Lst)
' - IsInList(W,Lst)             - определяет есть ли слово из текста W в списке Lst
' Т testFindInDIC()             - отладка FindInDIC(W,Dic)
' - IsInDIC(W,Dic)              - определяет есть ли слово W в словаре Dic
' - IsMatchList(W, DicList)     - возвращает TRUE если W входит в DicList
' T test ISML                   - отладка IsMatchList

Option Explicit

' Константы - используемых цветов
    Public Const rgbADSK = 2162853  ' характерный красно-коричневый цвет для Autodesk
    Public Const LimeG = 3329330    ' RGB(50, 205, 50)  лимонно-зеленый
    Public Const Antique = 11583693 ' RGB(205, 192, 176) Антик - серо-коричневый
    Public Const Gray = 8750469     ' серый
    
    Public Const Log = "Log"        ' Log лист
Function GetRep(RepName) As TOCmatch
'
' - GetRep(RepName) - находит и проверяет штамп отчета RepName
'   26.7.12
'    2.8.12 - NOP по пустому RepName
'   12.8.12 - StampR допускает альтернативное положение Штампа, например, "4, 1"
'   17.8.12 - FatalRep в отдельной подпрограмме; Activate RepName
'    9.9.12 - запись в Log только в match.xlsm; отладка записи Pass DBs; EOL для sfdc.xlsm
'   21.9.12 - отладка логики работы с match_environment при перемещении DirDBs
'   27.10.12 - работа с "голубыми" листами в TOCmatch

    Dim i As Long, EOL_TOC As Long
    Const TOClineN = 4  ' номер строки в TOCmatch описывающей саму себя
    
    If RepName = "" Then Exit Function
    
    If DB_MATCH Is Nothing Then
        Set DB_MATCH = FileOpen(F_MATCH)
        EOL_TOC = EOL(TOC, DB_MATCH)
        DB_MATCH.Sheets(TOC).Cells(TOClineN, TOC_EOL_COL) = EOL_TOC
'?'        GetRep = GetRep(TOC)        ' для TOCmatch - РЕКУРСИЯ для проверки штампа
    Else
        EOL_TOC = DB_MATCH.Sheets(TOC).Cells(TOClineN, TOC_EOL_COL)
    End If
            
    DirDBs = DB_MATCH.Path & "\"
    If DB_MATCH.Sheets(TOC).Cells(1, TOC_F_DIR_COL) <> DirDBs Then
        Dim Respond As Integer
        Respond = MsgBox("Файл <match.xlsx> загружен из необычного места:" _
            & vbCrLf & vbCrLf & "'" & DirDBs & "'" _
            & vbCrLf & vbCrLf & "Это теперь каталог файлов DBs? ", vbYesNo)
        If Respond <> vbYes Then End
        
'** новый DirDBs запишем в TOCmatch и во вспомогательный файл
        DB_MATCH.Sheets(TOC).Cells(1, TOC_F_DIR_COL) = DirDBs
        Dim F_match_env As Workbook ' вспомогательный файл c DirDBs
            ' при этом все отчеты из TOCmatch должны быть доступны!
        Dim rf As String
        For i = 8 To EOL_TOC
            rf = DB_MATCH.Sheets(TOC).Cells(i, TOC_REPFILE_COL)
            If rf <> "" Then FileOpen rf
        Next i
        
        Set F_match_env = Workbooks.Open(F_match_environment)
        With F_match_env.Sheets(1)
            .Cells(1, 1) = Now
            .Cells(1, 2) = DirDBs
        End With
        F_match_env.Close
'''''        Exit Function
    End If
    
    With DB_MATCH.Sheets(TOC)
'''''        For i = 4 To EOL(TOC, DB_MATCH)
'''''        For i = 4 To 177
        For i = TOClineN To EOL_TOC
            If .Cells(i, TOC_REPNAME_COL) = RepName Then GoTo FoundRep
        Next i
        FatalRep "GetRep ", RepName

FoundRep:
        RepTOC.Dat = .Cells(i, TOC_DATE_COL)
        RepTOC.Name = .Cells(i, TOC_REPNAME_COL)
        RepTOC.MyCol = .Cells(i, TOC_MYCOL_COL)
        RepTOC.ResLines = .Cells(i, TOC_RESLINES_COL)
        RepTOC.Made = .Cells(i, TOC_MADE_COL)
        RepTOC.RepFile = .Cells(i, TOC_REPFILE_COL)
        RepTOC.SheetN = .Cells(i, TOC_SHEETN_COL)
        RepTOC.EOL = .Cells(i, TOC_EOL_COL)
        RepTOC.CreateDat = .Cells(i, TOC_CREATED_COL)
        RepTOC.FormName = .Cells(i, TOC_FORMNAME)
    End With
    
'---- проверка штампа ----------
    Dim Str As Long, StC As Long
    Dim TestedStamp As String
    With RepTOC
        Select Case .RepFile
        Case F_MATCH:
            RepMatch = RepTOC
        Case F_1C:
            Set DB_1C = FileOpen(.RepFile)
            Rep1C = RepTOC
        Case F_SFDC:
            Set DB_SFDC = FileOpen(.RepFile)
            RepSF = RepTOC
        Case F_ADSK:
            Set DB_ADSK = FileOpen(.RepFile)
            RepADSK = RepTOC
        Case F_STOCK:
            Set DB_STOCK = FileOpen(.RepFile)
            RepStock = RepTOC
        Case F_TMP:
            Set DB_TMP = FileOpen(.RepFile)
''            RepWP = RepTOC
        Case Else: FatalRep "GetRep: файл штампа=" & .RepFile, RepName
        End Select
            
        CheckStamp i
        
        GetRep = RepTOC
'''        Workbooks(.RepFile).Sheets(.SheetN).Activate
    End With
End Function
Sub FatalRep(SubName, RepName)
'
' - FatalRep(SubName, RepName) - сообщение о фатальной ошибке при запросе RepName
' 17.8.12
' 9.8.12 -- более ясная диагностика по не найденному Штампу

    ErrMsg FATAL_ERR, SubName & "> Не найден Штамп в Документе " & RepName _
        & vbCrLf & vbCrLf & "Этот Документ надо загрузить в match заново!"
    Stop
'    End
End Sub
Function CheckStamp(iTOC As Long, _
    Optional NewRep As String = "", Optional NewRepEOL, Optional IsSF, _
    Optional InSheetN As Integer = 1) As Boolean
'
' - CheckStamp(iTOC) - проверка штампа в строке iTOC списка Документов в TOCmatch
' 15.8.2012
' 18.8.12 - CheckStamp оформлена как Bolean Function для использования в MoveToMatch
'           Optional параметры используются только для MoveToMatch
' 25.8.12 - входной Документ может находиться в листе InSheetN нового загружаемого файла
' 27.10.12 - помимо типов Штампа "=" и "I", введено "N" - Штамп не проверять

    Dim SR() As String, SC() As String
    Dim Str As Long, StC As Long
    
    Dim RepName As String
    Dim txt As String, TestedStamp As String
    Dim Typ As String
    Dim Continued As String
    Dim i As Long, j As Long
    
    CheckStamp = True
    
    With DB_MATCH.Sheets(TOC)
        SR = split(.Cells(iTOC, TOC_STAMP_R_COL), ",")
        SC = split(.Cells(iTOC, TOC_STAMP_C_COL), ",")
        txt = .Cells(iTOC, TOC_STAMP_COL)
        Typ = .Cells(iTOC, TOC_STAMP_TYPE_COL)
        If Typ = "N" Then GoTo Ex
        RepName = .Cells(iTOC, TOC_REPNAME_COL)
        Continued = .Cells(iTOC, TOC_PARCHECK_COL)
    End With
    
    With RepTOC
        For i = LBound(SR) To UBound(SR)
            For j = LBound(SC) To UBound(SC)
                Str = SR(i)
                StC = SC(j)
                If NewRep = "" Then
                    If .RepFile = F_SFDC Then Str = Str + .EOL
                    TestedStamp = Workbooks(.RepFile).Sheets(.SheetN).Cells(Str, StC)
                ElseIf IsMissing(IsSF) Then
                    Str = Str + NewRepEOL - SFresLines
                    TestedStamp = Workbooks(NewRep).Sheets(InSheetN).Cells(Str, StC)
                Else
                    If IsSF Then Str = Str + NewRepEOL - SFresLines
                    TestedStamp = Workbooks(NewRep).Sheets(InSheetN).Cells(Str, StC)
                End If
                If Typ = "=" Then
                    If txt <> TestedStamp Then GoTo NxtChk
                ElseIf Typ = "I" Then
                    If InStr(LCase$(TestedStamp), LCase$(txt)) = 0 Then GoTo NxtChk
                Else
                    ErrMsg FATAL_ERR, "Сбой в структоре TOCmatch: тип штампа =" & Typ
                End If
            
                If Continued <> "" Then CheckStamp iTOC + 1, NewRep, NewRepEOL, IsSF, InSheetN
Ex:             Exit Function
NxtChk:
            Next j
        Next i
        If NewRep = "" Then FatalRep "GetRep.CheckStamp", RepName
        CheckStamp = False
    End With
End Function
Function FileOpen(RepFile) As Workbook
'
' - FileOpen(RepFile)   - проверяет, открыт ли RepFile, если нет - открывает
'   26.7.12
    
    Dim W As Workbook
    For Each W In Application.Workbooks
        If W.Name = RepFile Then
            Set FileOpen = W
            Exit Function
        End If
    Next W
    
    If DirDBs = "" Then
        Dim F_match_env As Workbook ' вспомогательный файл c DirDBs
        Set F_match_env = Workbooks.Open(F_match_environment)
        DirDBs = F_match_env.Sheets(1).Cells(2, 1)
        F_match_env.Close
    End If
    
    Set FileOpen = Workbooks.Open(DirDBs & RepFile, UpdateLinks:=False)
End Function
Sub WrTOC()
'
' - WrTOC() - записывает структуру Public RepTOC в оглавление match.Sheets(TOC)
'   5.8.2012
'  12.8.12 - "серые" колонки описывающие Штамп не записываем
'  17.8.12 - еще ряд полей не записывыем в match.xlsm и использование FatalRep
'   2.9.12 - дополнительные ограничения записи в TOCmatch
' 28.10.12 - записывает в TOCmatch дату создания CreateDat

    Dim i As Long
    Const BEGIN = 8 ' начало списка обрабатываемых Документов
    
    If RepTOC.Name = "" Then FatalRep "WrTOC", "<пусто>"
    For i = BEGIN To BIG
        If DB_MATCH.Sheets(1).Cells(i, TOC_REPNAME_COL) = RepTOC.Name Then GoTo FoundRep
    Next i
    FatalRep "WrTOC", RepTOC.SheetN

FoundRep:
    With DB_MATCH.Sheets(TOC)
        .Cells(i, TOC_DATE_COL) = RepTOC.Dat
'''        .Cells(i, TOC_REPNAME_COL) = RepTOC.Name
        .Cells(i, TOC_EOL_COL) = RepTOC.EOL
'''        .Cells(i, TOC_MYCOL_COL) = RepTOC.MyCol
'''        .Cells(i, TOC_RESLINES_COL) = RepTOC.ResLines
        .Cells(i, TOC_MADE_COL) = RepTOC.Made
'''        .Cells(i, TOC_REPFILE_COL) = RepTOC.RepFile
'''        .Cells(i, TOC_SHEETN_COL) = RepTOC.SheetN
'''        .Cells(i, TOC_STAMP_COL) = RepTOC.Stamp
'''        .Cells(i, TOC_STAMP_TYPE_COL) = RepTOC.StampType
'''        .Cells(i, TOC_STAMP_R_COL) = RepTOC.StampR
'''        .Cells(i, TOC_STAMP_C_COL) = RepTOC.StampC
        .Cells(i, TOC_CREATED_COL) = RepTOC.CreateDat
'''        .Cells(i, TOC_PARCHECK_COL) = RepTOC.ParChech
'''        .Cells(i, TOC_REPLOADER_COL) = RepTOC.Loader
        .Cells(1, 1) = Now
    End With
End Sub
Sub InsMyCol(F As String, Optional FS As String = "")
'
' S InsMyCol(F [,FS]) - вставляем колонки в лист слева по шаблону F и пятку из FS
'
'   * Если заголовок колонки шаблона пятки пустой - пропускаем
'   * Если в строке 2 шапки шаблона "V" - переписываем шапку из шаблона
'
'  10.8.12
'  15.8.12 - Optional FS
'  26.8.12 - RowHeight шапки как в шаблоне; если строке 2 "V" - копируем шапку
'  31.8.12 - внедрение StepIn
'  11.9.12 - перенос форм в Headers файла match.xlsm
'  1.10.12 - копирование заголовка колонки в Шапку по COPY_HDR в строке 2 Шаблона
'  4.11.12 - использование R=GetRep(ActiveSheet.Name)
' 19.11.12 - COPY_HDR - copy Шаблона вместо присваивания текстового содержимого

    Const COPY_HDR = "CopyHdr"

    StepIn
    
    Dim R As TOCmatch   'R - структура TOCmatch для SFD
    Dim FF As Range
    Dim i As Integer
    Set FF = DB_MATCH.Sheets(Header).Range(F)
    
    R = GetRep(ActiveSheet.Name)
    With Workbooks(R.RepFile).Sheets(R.SheetN)
'---- А может мы уже эту колонку вставляли?
        If .Cells(1, 1) = FF.Cells(1, 1) Then Exit Sub

'---- вставляем колонки по числу MyCol
        For i = 1 To R.MyCol
            .Cells(1, 1).EntireColumn.Insert
        Next i
'---- задаем ширину и заголовки вставленных колонок
        For i = 1 To FF.Columns.Count
            .Columns(i).ColumnWidth = FF.Cells(3, i)
            If FF.Cells(2, i) = COPY_HDR Then
                FF.Cells(1, i).Copy Destination:=.Cells(1, i)
            End If
        Next i
'---- копируем колонки MyCol от верха до EOL
        For i = 1 To R.MyCol
            FF.Cells(1, i).Copy Destination:=.Cells(1, i)
            FF.Cells(2, i).Copy Destination:=.Cells(2, i)
        Next i
        .Rows(1).RowHeight = FF.Rows(1).RowHeight
        .Range(.Cells(2, 1), .Cells(R.EOL, R.MyCol)).FillDown
'---- вставляем пятку по шаблону в FS
        If FS = "" Then Exit Sub
        Set FF = DB_MATCH.Sheets(Header).Range(FS)
        For i = 1 To FF.Columns.Count
            If FF.Cells(1, i) <> "" Then
                FF.Columns(i).Copy Destination:=.Cells( _
                    R.EOL + R.ResLines - FF.Rows.Count + 1, i)
            End If
        Next i
    End With
End Sub
Sub MS(Msg)
'
'   - MS(Msg)- вывод сообщения на экран и в LogWr
'   11.6.12
    ErrMsg TYPE_ERR, Msg
End Sub

Sub ErrMsg(ErrMode, Msg)
'
' - ErrMsg(ErrMode, MSG) - вывод сообщения об ощибке в Log и на экран
'                          Коды ErrMode определены в Declaration
'   31.5.12

    Dim ErrType As String, Respond As String

    Select Case ErrMode
    Case WARNING:
        LogWr "< WARNING > " & Msg
        Exit Sub
        
    Case TYPE_ERR:
        LogWr "ВНИМАНИЕ:" & Msg
        Respond = MsgBox(Msg & vbCrLf & vbCrLf & "Продолжить?", vbYesNo)
        If Respond = vbNo Then
            ExRespond = False
            Stop
        End If
        Exit Sub
        
    Case FATAL_ERR:
Fatal:  ErrType = "<! ERROR !> "
        LogWr ErrType & Msg
        MsgBox Msg, , ErrType
        Stop
        Exit Sub
    Case Else:
        ErrMode = FATAL_ERR
        GoTo Fatal
    End Select
End Sub
Sub LogWr(Msg)
'
' запись сообщения msg в Log-лист
'   15.2.2012
'   26.6.12 - match 2.0
'    9.9.12 - вместо множества Log в разных файлах вседа пишем в Log match.xlsm

    Dim N   ' номер строки в Log
    
    With DB_MATCH.Sheets(Log)
        N = .Cells(1, 4)
        N = N + 1
        .Cells(N, 1) = Date
        .Cells(N, 2) = Time
        .Cells(N, 3) = Msg
        .Cells(1, 4) = N
    End With
End Sub
Sub LogReset()
'
' сброс Log листа и его очистка
' запускается кнопкой [Reset] на лосте Log
'   19.2.2012

    Dim N
    Sheets(Log).Select
    Cells(1, 3).Select
    N = Cells(1, 4)
    Cells(1, 4) = 1
    Rows("2:" & N).Delete
    LogWr ("LogReset")
End Sub
Function AutoFilterReset(SheetN) As Integer
'
' подпрограмма сброса и взвода фильтра листа SheetN в первой строке
'      возвращает количество строк в SheetN
' 16.1.2012

    Sheets(SheetN).Select
    ActiveSheet.AutoFilterMode = False  ' собственно сброс фильтра
    ActiveWindow.FreezePanes = False    ' Top Row Freeze
    Rows("1:1").AutoFilter              ' включаем/выключаем AutoFilter
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    AutoFilterReset = Sheets(SheetN).UsedRange.Rows.Count
    Range("A" & AutoFilterReset).Activate ' выбираем ячейку внизу листа
End Function
Sub SheetsCtrlH(SheetN, FromStr, ToStr)
'
' подпрограмма замены (Ctrl/H) строки FromStr на ToStr в листе SheetN
'   7.1.2012
'  27.1.2012 - убрал Activate/Select
    
    Call AutoFilterReset(SheetN)

    Cells.Replace What:=FromStr, Replacement:=ToStr, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
Sub PerCent(Row, Col)
'
' Подпрограмма форматирования %
'   26.1.12

    Cells(Row, Col).NumberFormat = "@"
    Cells(Row, Col) = Cells(Row, Col) & "%"
End Sub
Sub CurCode(Row, Col, CurCol)
'
' подпрограмма форматирования числа в ячейке (Row,Col).
' Код валюты в той же строке в колонке Col
'   20.1.12

    Select Case Cells(Row, CurCol)
        Case "RUB"
            Cells(Row, Col).NumberFormat = "_-* #,##0.00[$р.-419]_-;-* #,##0.00[$р.-419]_-;_-* ""-""??[$р.-419]_-;_-@_-"
        Case "EUR"
            Cells(Row, Col).NumberFormat = "_-[$€-2] * #,##0.00_ ;_-[$€-2] * -#,##0.00 ;_-[$€-2] * ""-""??_ ;_-@_ "
        Case "USD"
            Cells(Row, Col).NumberFormat = "_-[$$-409]* #,##0.00_ ;_-[$$-409]* -#,##0.00 ;_-[$$-409]* ""-""??_ ;_-@_ "
        Case Else
            MsgBox "ERROR in WPopp: Странный код валюты = " & Cells(Row, CurCol), , "ERROR!"
    End Select
End Sub
Function CurRate(Cur) As Double
'
' возвращает число - курс к рублю по коду валюты Cur по таблице Currence на листе We
'   21.2.2012
'   20.8.12 - распознавание "руб"
'    4.9.12 - адресация Sheets(We)

    Dim S

    CurRate = 1
    If InStr(LCase(Cur), "руб") > 0 Or Trim(Cur) = "" Then Exit Function

    On Error GoTo Col2
    S = WorksheetFunction.VLookup(Cur, DB_MATCH.Sheets(We).Range("RUB_Rate"), 3, False)
    On Error GoTo 0
    GoTo Convert
Col2:
    On Error GoTo 0
    S = WorksheetFunction.VLookup(Cur, DB_MATCH.Sheets(We).Range("RUB_Rate_2"), 2, False)
Convert:
    CurRate = Replace(S, ".", ",")
End Function
Function CurISO(Cur1C)
'
' возвращает код валюты в стандарте ISO, преобразовав его из вида 1С
'   18.3.2012
'    4.9.2012 - адресация Sheets(We)
'   19.9.12 - по умолчанию CurISO="RUB"

    CurISO = "RUB"
    On Error Resume Next
    CurISO = WorksheetFunction.VLookup(Cur1C, DB_MATCH.Sheets(We).Range("Currency"), 2, False)
    On Error GoTo 0
End Function
Function DDMMYYYY(D) As String
'
' преобразование даты d в текстовый формат DDMMYYYY
'   14.2.2012
    DDMMYYYY = Day(D) & "." & Month(D) & "." & Year(D)
End Function
Function Dec(A) As String
'
' преобразование числа а в текстовый формат с десятичной точкой
'   14.2.2012

    Dec = "'" & WorksheetFunction.Substitute(A, ",", ".")
End Function
Sub testEOL()
    Dim A, b, C
    A = EOL(1)
    b = EOL(2)
        Dim F As Workbook
        Set F = Workbooks.Open(F_SFDC, UpdateLinks:=True, ReadOnly:=True)
    C = EOL(1, F)
        F.Close SaveChanges:=False
End Sub
Function EOL(ByVal SheetN As String, Optional F As Workbook = Nothing)
'
' - EOL(SheetN,[F]) - возвращает количество строк в листе SheetN файла F
'   20/1/2012
'   4/2/2012 - обработка ошибок On Error
'   20/2/2012 - внедрение Option Explicit
'   12.5.12 - Sheets(SheetN).Select исключен
'   24.6.12 - AllCol - Public
'   29.6.12 - match 2.0 - открыть файл S, если указан параметр
'   31.7.12 - файл ActiveWorkbook а не ThisWorkbook по умолчанию,
'             ByVal SheetN As String - преобразование передаче параметра
'   20.8.12 - по отсутствующему SheetN возвращается EOL = -1

    Dim i

    If F Is Nothing Then
        Set F = ActiveWorkbook
    End If
    
    EOL = -1
    On Error Resume Next
    EOL = F.Sheets(SheetN).UsedRange.Rows.Count
    On Error GoTo 0
    If EOL <= 0 Then Exit Function
    
    With F.Sheets(SheetN)
        AllCol = .UsedRange.Columns.Count
        Do
            For i = 1 To AllCol
                If .Cells(EOL, i) <> "" Then Exit Do
            Next i
            If EOL <= 1 Then Exit Do
            EOL = EOL - 1       ' иногда UsedRange оставляет пустые строки,
        Loop                    '   .. например, если в строке есть невидимый формат
    End With
End Function
Sub RowDel(RowStr As String)
'
' - RowDel(RowStr) - удаляет строки активного листа в соответствии с RowStr
'   25.8.12
    StepIn
    ActiveSheet.Rows(RowStr).Delete
End Sub
Function CSmatch(Val, Col) As Double
'
' - CSmatch(Val,Col) - Case Sensitive match возвращает номер строки с Val в колонке Col.
'                   Если Val не найден- возвращает 0. Лист для поиска Val должен быть Selected.
' 8/7/12

    Dim CheckCS
    Dim N As Long
    N = 1
    Do
        CSmatch = 0
        On Error Resume Next
        CSmatch = Application.Match(Val, Range(Cells(N, Col), Cells(BIG, Col)), 0) + N - 1
        CheckCS = Cells(CSmatch, Col)
        On Error GoTo 0
        If IsEmpty(CSmatch) Or Not IsNumeric(CSmatch) Or CSmatch <= 0 Then Exit Function
        N = CSmatch + 1
    Loop While Val <> CheckCS
End Function
Function CSmatchSht(Val, Col, Sht, Optional ByVal FromN As Long = 1) As Long
'
' - CSmatch(Val,Col,Sht) - Case Sensitive match возвращает номер строки с Val в колонке Col.
'                   Если Val не найден- возвращает 0. Sht - лист для поиска Val.
' 27.9.12
' 25.11.12 - Optional FromN

    Dim CheckCS
''    Dim N As Long
''    N = 1
    Do
        CSmatchSht = 0
        On Error Resume Next
        CSmatchSht = Application.Match(Val, Range(Sht.Cells(N, Col), Sht.Cells(BIG, Col)), 0) _
            + FromN - 1
        CheckCS = Sht.Cells(CSmatchSht, Col)
        On Error GoTo 0
''        If IsEmpty(CSmatchSht) Or Not IsNumeric(CSmatchSht) Or CSmatchSht <= 0 Then Exit Function
        If Not IsNumeric(CSmatchSht) Or CSmatchSht <= 0 Then Exit Function
        FromN = CSmatchSht + 1
    Loop While Val <> CheckCS
End Function
Sub ClearSheet(SheetN, HDR_Range As Range)
'
' Полная очистка SheetN и перенос в него заголовка из листа Нeader.HDR_Range
'   4.2.2012
'  11.2.2012 - пересмотр спецификаций
'  10.3.12 - изменение спецификации - параметр HRD_Range
'  25.3.12 - листы NewContract и NewContractLnk
'  17.4.12 - лист A_Acc - новые Организации
'  18.4.12 - лист A_Dic - Словарь Организаций
'  28.4.12 - лист NewOrderList - лист Новых Заказов
'  13.5.12 - лист P_ADSKlink - новые связки Платеж - ADSK
'  15.5.12 - лист SF_PA связей Платежей с Контрактами ADSK
'   6.6.12 - Delete старый лист, создаем новый
'  11.6.12 - листы A_Acc и AccntUpd
'  12.6.12 - лист BTO_SHEET - лог для писем БТО

    DB_MATCH.Sheets(SheetN).Activate
    
' -- стираем старый лист
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(SheetN).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
' -- создаем новый лист
    Sheets.Add After:=Sheets(Sheets.Count)  ' создаем новый лист в конце справа
    ActiveSheet.Name = SheetN
    ActiveSheet.Tab.Color = RGB(50, 153, 204)   ' Tab голубой
   
    HDR_Range.Copy Sheets(SheetN).Cells(1, 1)   ' копируем шапку из Header
    
    Select Case SheetN
    Case O_NewOpp:      EOL_NewOpp = 1
    Case P_Paid:        EOL_NewPay = 1
    Case NewContract:       EOL_NewContr = 1
    Case NewContractLnk:    EOL_ContrLnk = 1
    Case P_PaymentUpd:  EOL_PaymentUpd = 1
    Case A_Dic:         EOL_DIC = 1
    Case A_Acc:         EOL_AdAcc = 1
    Case AccntUpd:      EOL_AccntUpd = 1
    Case NewOrderList:  EOL_NewOrderList = 1
    Case P_ADSKlink:    EOL_ADSKlnkPay = 1
    Case SF_PA:         EOL_SFlnkADSK = 1
    Case NewSN:         EOL_NewSN = 1
    Case BTO_SHEET:     EOL_BTO = 1
'    Case WP:        EOL_WP = 1
    Case Else
        MsgBox "ClearSheet: Очистка странного листа '" & SheetN & "'" _
            , , "ERROR!"
        Stop
    End Select
End Sub
Sub SheetSort(SheetN, Col)
'
' Сортируем лист SheetN по колонке Col
'   22.1.2012
'   21.2.2012 - Option Explicit
'   19.4.12 - AutoFilterReset

    Dim Name As String

'    Sheets(SheetN).Select
    Call AutoFilterReset(SheetN)

    Name = ActiveSheet.Name
    
    With ActiveWorkbook.Worksheets(Name).AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add key:=Cells(1, Col), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub SheetDedup(SheetN, Col)
'
' Удаляем строки - дубликаты в лист SheetN по колонке Col,
'                  выполнив сортировку по этой колонке
'   19.4.2012

    Dim i, prev, X, EOL_SheetN As Integer
    
    Call SheetSort(SheetN, Col)
    EOL_SheetN = EOL(SheetN)
    
    prev = "": i = 2
    Do
        X = Sheets(SheetN).Cells(i, Col)
        If X = prev Then
            Rows(i).Delete
            EOL_SheetN = EOL_SheetN - 1
        Else
            i = i + 1
            prev = X
        End If
    Loop While i < EOL_SheetN
End Sub
Sub SheetDedup2(SheetN, ColSort, СolAcc, ColIdSF)
'
' - SheetDedup2(SheetN, ColSort, ColAcc, ColIdSF)  - сортируем лист SheetN
'          по колонке SortCol затем значения ColAcc и ColIdSF сливем с "+"
'   23.5.2012
'   23.11.12 - отладка в match2.0

    Dim i As Integer, EOL_SheetN As Integer
    Dim prev As String, X As String
    Dim PrevAcc As String, NewAcc As String
    Dim PrevSFid As String, NewSFid As String
    
    Call SheetSort(SheetN, ColSort)
    EOL_SheetN = EOL(SheetN)
    
    prev = "": i = 2
    With Sheets(SheetN)
        Do
            X = .Cells(i, ColSort)
            If X = prev Then
                PrevAcc = .Cells(i - 1, СolAcc)
                PrevSFid = .Cells(i - 1, ColIdSF)
                NewAcc = .Cells(i, СolAcc)
                NewSFid = .Cells(i, ColIdSF)
                If PrevAcc <> "" And NewAcc <> "" And PrevAcc <> NewAcc Then
                    PrevAcc = PrevAcc & "+" & NewAcc
                ElseIf PrevAcc = "" And NewAcc <> "" Then
                    PrevAcc = NewAcc
'                ElseIf PrevAcc <> "" And NewAcc = "" Then
'                ElseIf PrevAcc = "" And NewAcc = "" Then
'                   в двух последних случаях ничего не делаем
                End If
                If PrevSFid <> "" And NewSFid <> "" And PrevSFid <> NewSFid Then
                    .Cells(i - 1, ColIdSF) = PrevSFid & "+" & NewSFid
                End If
                .Cells(i - 1, СolAcc) = PrevAcc
                Rows(i).Delete
                EOL_SheetN = EOL_SheetN - 1
            Else
                i = i + 1
                prev = X
            End If
        Loop While i < EOL_SheetN
    End With
End Sub
Sub DateCol(ByVal SheetN As String, ByVal Col As Integer)
'
' преобразование колонки Col в листе SheetN из текста вида DD.MM.YY в формат Date
'   20.4.12
'   3.10.12 - GetRep вместо EOL

    Dim i, dd, MM, YY As Integer
    Dim Dat As Date
    Dim D() As String
    
    Dim R As TOCmatch
    R = GetRep(SheetN)
    
    For i = 1 To R.EOL
        D = split(Sheets(SheetN).Cells(i, Col), ".")
        If UBound(D) = 2 Then
            dd = D(0)
            If dd < 1 Or dd > 31 Then GoTo Nxt
            MM = D(1)
            If MM < 1 Or MM > 12 Then GoTo Nxt
            YY = D(2)
            Dat = dd & "." & MM & "." & YY
            Sheets(SheetN).Cells(i, Col) = Dat
        End If
Nxt:
    Next i
End Sub
Sub DateSort(ByVal SheetN As String, ByVal Col As Integer)
'
' S DateSort(SheetN, Col) - преобразование колонки Col из текстового формата в Date
'                           и сортировка по этой колонке от старых к новым датам
'   31.7.12
'   31.8.12 - оформлен как Step со StepIn

    StepIn
'''    Sheets(SheetN).Activate
    DateCol SheetN, Col
    SheetSort SheetN, Col
End Sub
Sub HideLns(FrR, ToR, Col, _
    Optional Criteria As String, Optional HideFlag As Boolean = True)
'
' скрывает (Hide) строки листа, удовлетворяющие Criteria
' просматриваются поля в колонке Col со строки FrR до ToR
' - Если параметр HideFlag=FALSE, то наоборот делает Ungide
' - Если Col = 0 - Hide или Unhide все строки от FrR до ToR
'   25.1.2012
'    4.2.2012 - Unhide режим по HideFlag=FALSE

    Dim RowsList As String      ' список строк для Hide
    Dim RowsListLng As Integer  ' длина списка строк
    
    RowsList = ""
    RowsListLng = 0
    If Col = 0 Then
       RowsList = FrR & ":" & ToR
       RowListLng = 1
    Else
        For i = FrR To ToR
            If Cells(i, Col) = Criteria Then
                If RowsListLng > 0 Then RowsList = RowsList & ","
                RowsList = RowsList & i & ":" & i
                RowsListLng = RowsListLng + 1
            End If
        Next i
        If RowsListLng = 0 Then Exit Sub    ' если ничего делать не надо
    End If
    Range(RowsList).Select      ' образец:  Range("30:30,31:31")
    Selection.EntireRow.Hidden = HideFlag
End Sub
Sub Progress(Pct)
'
' вывод Progress Bar'а - процентов выполнения кода
'   15/2/2012
'   26.5.12 - MsgBox каждые 20 сек во время исполнения Progress
'   29.5.12 - изменение формы
'    7.8.12 - попытка заменить Форму на StatusBar
'   31.8.12 - внедрение StepIn
            
    Application.StatusBar = PublicProcName & "> " _
        & "Шаг " & PublicStepName _
        & ": " & Format(Pct, "0%")
    End Sub
Sub StopSub()
'
' StopSub() аварийное завершение процесса - вызывается по Событию FATAL ERROR
'
    
    MsgBox "Аварийная остановка StopSub", , "FATAL ERROR"
    Stop
End Sub
Function RemDelimiters(S)
'
' RemDelimeters(S) - замена спец символов в строке S на пробелы
' 19.4.12 А.Пасс

    Dim smb, i

    For i = 1 To Len(S)
        
        ' допускаем цифры, латинские или русские буквы
        smb = Mid(S, i, 1)
        If Not ((smb >= "0" And smb <= "9") _
                Or (smb >= "A" And smb <= "Z") _
                Or (smb >= "a" And smb <= "z") _
                Or (smb >= "А" And smb <= "Я") _
                Or (smb >= "а" And smb <= "я") _
                Or smb = "ъ" Or smb = "Ъ" _
                Or smb = "ё" Or smb = "Ё") Then
            Mid(S, i, 1) = " "      ' символ недопустим, заменяем на пробел
        End If
    Next i
    RemDelimiters = S
End Function
Function Compressor(S1 As Variant)
'
' удаление лишних пробелов внутри строки
'   7.3.12  из Интернет
'   7.6.12 удаляем vbCcLf

    Dim S As Variant
    S = Replace(S1, vbCrLf, " ")
    S = Trim(S)
    While InStr(1, S, "  ") <> 0
       S = Left(S, InStr(1, S, "  ") - 1) & Right(S, Len(S) - InStr(1, S, "  "))
    Wend
    Compressor = S
End Function
Sub testFindInLst()
'
' Т testFindInLst() - отладка FindInLst(W,Lst)
'   24/5/12
    Dim A
    A = FindInLst("Autodesk Plant Design Suite Premium 2012 New SLM", "DIC_GoodADSK")
End Sub
Function FindInLst(W, Lst) As String
'
' - IsInList(W,Lst)  - определяет есть ли слово из текста W в списке Lst
'   24.5.12

    Dim S As Range
    Dim lW As String, V As String
    
    lW = LCase$(W)
    For Each S In Range(Lst).Rows
        V = S.Cells(1, 1)
        If InStr(lW, LCase$(V)) <> 0 Then
            FindInLst = V
            Exit For
        End If
    Next S
End Function
Sub testFindInDIC()
'
' Т testFindInDIC() - отладка FindInDIC(W,Dic)
'   7/5/12
    Dim A
    A = FindInDIC("картридж", "Goods")
End Sub
Function FindInDIC(W, Dic) As String
'
' - IsInDIC(W,Dic)  - определяет есть ли слово W в словаре Dic
'       распознавание происходит по второй колонке Dic
'       в ней слова для распознавания, разделенные запятыми.
'       если во второй колонке пусто - сравниваем со словом в первой
'   7.5.12

    FindInDIC = ""
    Dim S As Range
    Dim ValWord As String
    Dim Article As String

    For Each S In Range(Dic).Rows
        ValWord = S.Cells(1, 1)
        Article = S.Cells(1, 2)
        If Article = "" Then Article = ValWord
        If IsMatchList(W, Article) Then
            FindInDIC = ValWord
            Exit For
        End If
    Next S
End Function
Function IsMatchList(W, DicList) As Boolean
'
' - IsMatchList(W, DicList) - возвращает TRUE если W входит в DicList
'   7.5.12

    IsMatchList = False
    If W = "" Or DicList = "" Then Exit Function
    
    Dim X() As String
    Dim i As Integer
    Dim lW As String
    
    lW = LCase$(W)
    X = split(DicList, ",")
    
    For i = LBound(X) To UBound(X)
        If InStr(lW, LCase$(X(i))) <> 0 Then
            IsMatchList = True
            Exit Function
        End If
    Next i
End Function
Sub testISML()
'
' T test ISML - отладка IsMatchList
' 7/5/12
    Dim A As Boolean
    A = IsMatchList("", "мышка,кошка,лев")
    A = IsMatchList("собака", "мышка,кошка,соб,лев")
    A = IsMatchList("собака", "мышка,кошка,лев")
End Sub
Sub ScreenUpdate(TurnOn As Boolean)
'
' - ScreenUpdate(ToDo) - switch off Screen Update if TurnOn = False
'   8.11.12

    If Not TurnOn Then
        With Application
            .ScreenUpdating = False
'            .EnableEvents = False
            .DisplayAlerts = False
        End With
    Else
        With Application
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
            .DisplayStatusBar = True
            .DisplayAlerts = True
        End With
    End If
End Sub
'''
''''?????????????????????????????????????????????????????????????????????????
''''?????????????????? процедуры, подлежащие удалению  ??????????????????????
''''?????????????????????????????????????????????????????????????????????????
'''Sub ModStart(Report)
''''
'''' - ModStart(Report)    - начало работы с отчетом Report, проверки и инициализации
''''
''''  26.7.12  - переписано для match 2.0
'''
'''    GetRep TOC
'''
'''    Select Case Report
'''    Case REP_1C_P_LOAD:
'''        Doing = "Загружаем новый отчет по Платежам 1С в базу 1C.xlsm"
'''        GetRep SF
'''        GetRep PAY_SHEET
'''''        CheckSheet PAY_SHEET
'''''        EOL_PaySheet = RepTOC.EOL
'''    Case REP_1C_P_PAINT:
'''        Doing = "Раскрашиваем лист Платежей базы 1C.xlsm"
'''    Case REP_1C_SFACCFIL:
'''        Doing = "Заполнение колонки 1 для листа Платежей"
'''        GetRep PAY_SHEET
'''        EOL_PaySheet = RepTOC.EOL
'''''''''''''''        EOL_SFacc = EOL(SFacc, F_SFDC) - SFresLines
'''    Case REP_SF_LOAD:
'''        Doing = "Загрузка Платежей из Salesforce - SF"
'''        Set DB_1C = Workbooks.Open(DirDBs & F_1C, UpdateLinks:=False, ReadOnly:=True)
'''        GetRep PAY_SHEET
'''        EOL_PaySheet = RepTOC.EOL
'''        GetRep SF
'''        EOL_SF = RepTOC.EOL
'''''        CheckSheet PAY_SHEET, 1, PAYDOC_COL, Stamp1Cpay1
'''''        CheckSheet PAY_SHEET, 1, PAYDATE_COL, Stamp1Cpay2
'''''        EOL_PaySheet = EOL(PAY_SHEET) - PAY_RESLINES
'''''        EOL_SFacc = EOL(SFacc, F_SFDC) - SFresLines
''''''        P = True
''''    Case REP_1C_С_LOAD:
''''    Case Acc1C:
'''    Case Else:
'''        ErrMsg FATAL_ERR, "Запрошен неизвестный отчет"
'''        End
'''    End Select
'''
''''''''''    With Application
''''''''''        .DisplayStatusBar = True
''''''''''        .StatusBar = Msg
''''''''''' для ускорения Excel отключаем вывод и др.
''''''''''        .ScreenUpdating = False
''''''''''        .Calculation = xlCalculationManual
''''''''''        .EnableEvents = False
''''''''''        .DisplayAlerts = False
''''''''''    End With
''''''''''    ActiveSheet.DisplayPageBreaks = False
''''''''''    Call AutoFilterReset(SheetN)
''''''''''
''''''''''' ---- определение EOL для всех основных листов
''''''''''    EOL_DogSheet = EOL(DOG_SHEET) - DOGRES
''''''''''    EOL_SF = EOL(SF) - SFresLines
''''''''''    EOL_SFD = EOL(SFD) - SFresLines
''''''''''    EOL_SFopp = EOL(SFopp) - SFresLines
''''''''''    EOL_SFacc = EOL(SFacc) - SFresLines
''''''''''    EOL_Acc1C = EOL(Acc1C) - ACC1C_RES
''''''''''    EOL_ADSKfrSF = EOL(ADSKfrSF) - SFresLines
''''''''''    EOL_Stock = EOL(STOCK_SHEET)
''''''''''    EOL_PaySheet = EOL(PAY_SHEET) - PAY_RESLINES
''''''''''    EOL_SFlnkADSK = EOL(SF_PA) - SFresLines
''''''''''
''''''''''    Select Case SheetN
''''''''''    Case PAY_SHEET:     ModStart = EOL_PaySheet
''''''''''    Case DOG_SHEET:     ModStart = EOL_DogSheet
''''''''''    Case Acc1C:         ModStart = EOL_Acc1C
''''''''''    Case STOCK_SHEET:   ModStart = EOL_Stock
''''''''''    Case SF:            ModStart = EOL_SF
''''''''''    Case SFD:           ModStart = EOL_SFD
''''''''''    Case SFacc:         ModStart = EOL_SFacc
''''''''''    Case SF_PA:         ModStart = EOL_SFlnkADSK
''''''''''    Case Else:
''''''''''        ModStart = EOL(SheetN)
''''''''''    End Select
''''''''''' ----
'''    ExRespond = True
'''
''''    Range("A1:A" & ModStart).EntireRow.Hidden = False
'''    With ProgressForm
'''        .Show vbModeless
'''        .ProgressLabel.Caption = Doing
'''    End With
'''    LogWr ""
'''    LogWr (Doing)
'''End Sub
''''''Sub ModEnd()
'''''''
''''''' - ModEnd() - подпрограмма завершения работы Модуля
'''''''  15.2.2012
'''''''  19.4.12  - восстановление вывода Excel
'''''''  2.7.12  - match 2.0
''''''' 20.7.12 - переписываем TOCmatch to RepTOC
''''''
''''''    WrTOC
''''''    Close
''''''
'''''''    i = AutoFilterReset(SheetN)
'''''''    ActiveSheet.Range("A" & i).Select
''''''    ProgressForm.Hide
''''''' восстанавливаем вывод Excel и др
''''''    With Application
''''''        .StatusBar = False
''''''        .ScreenUpdating = True
''''''        .Calculation = xlCalculationAutomatic
''''''        .EnableEvents = True
''''''        .DisplayStatusBar = True
''''''        .DisplayAlerts = True
''''''    End With
''''''    ActiveSheet.DisplayPageBreaks = True
''''''    LogWr (Doing & " - ГОТОВО!")
''''''End Sub
''''''
