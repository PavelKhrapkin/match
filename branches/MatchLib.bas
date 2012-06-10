Attribute VB_Name = "MatchLib"
'---------------------------------------------------------------------------
' Библиотека подпрограмм проекта MatchSF-1C
'
' П.Л.Храпкин, А.Пасс 7.6.2012
'
' - ModStart(SheetN, MSG)       - начало модуля работы с Листом SheetN
' - PublicVarInit()             - инициализация глобальных переменных EOL и др
' - ModEnd(SheetN)              - завершение Модуля, работающего с листом SheetN
' - ErrMsg(ErrMode, MSG)        - вывод сообщения об ощибке в Log и на экран
' - LogWr(msg)                  - запись сообщения msg в Log list
'(*)LogReset()                  - сброс и очистка Log листа
' - ActiveFilterReset(SheetN)   - сброс и активизация автофильтра листа SheetN
' - SheetsCtrlH(SheetN, FromStr, ToStr) - замена текста FromStr на ToStr
'                                 в листе SheetN
' - Pnt(SheetN,Col,Criteria,Color,Mode) - окраска на SheetN цветом Color по фильтру
' - PerCent(Row, Col)           - форматирование ячейки (Row,Col) с процентами
' - CurCode(Row, Col, CurCol)   - формат ячейки (Row,Col) по коду валюты в CurCol
' - CurRate(Cur)                - возвращает курс валюты к рублю по коду Cur для We
' - CurISO(Cur1C)               - возвращает код валюты в стандарте ISO
' - DDMMYYYY(d)                 - преобразование даты d в текстовый формат DDMMYYYY
' - Dec(a)                      - формат числа а в виде текста с десятичной точкой
' - EOL(SheetN)                 - возвращает номер последней строки листа SheetN
' - ClearSheet(SheetN, HDR_Range) - очистка листа SheetN и запись в него шапки
' - SheetSort(SheetN, Col)      - сортировка листа SheetN по колонке Col
' - SheetDedup(SheetN, Col)     - cортировка и дедупликация SheetN по колонке Col
' - SheetDedup2(SheetN, ColSort,ColAcc) - сортировка и слияние листа SheetN
'                                 по колонкам ColSort, ColAcc
' - DateCol(SheetN, Col)        - преобразование колонки Col из текста в Дату
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

Function ModStart(SheetN, MSG, Optional P As Boolean = False) As Integer
'
' подпрограмма - начало работы Модуля с листом SheetN, возвращает количество строк
'  23.3.2012
'  19.4.12  - отключение вывода Excel

    Call PublicVarInit
    Doing = MSG
    With Application
        .DisplayStatusBar = True
        .StatusBar = MSG
' для ускорения Excel отключаем вывод и др.
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    ActiveSheet.DisplayPageBreaks = False
    Call AutoFilterReset(SheetN)
    ModStart = EOL(SheetN)
    Range("A1:A" & ModStart).EntireRow.Hidden = False
    If P Then
        With ProgressForm
            .Show vbModeless
            .ProgressLabel.Caption = Doing
        End With
    End If
    LogWr ""
    LogWr (Doing)
End Function
Sub PublicVarInit()
'
' PublicVarInit()   - инициализация глобальных переменных
'   15.5.12

    EOL_DogSheet = EOL(DOG_SHEET) - DOGRES
    EOL_SFD = EOL(SFD) - SFresLines
    EOL_SFopp = EOL(SFopp) - SFresLines
    EOL_SFacc = EOL(SFacc) - SFresLines
    EOL_Acc1C = EOL(Acc1C) - ACC1C_RES
    EOL_ADSKfrSF = EOL(ADSKfrSF) - SFresLines
    EOL_Stock = EOL(STOCK_SHEET)
    EOL_PaySheet = EOL(PAY_SHEET) - PAY_RESLINES
    EOL_SFlnkADSK = EOL(SF_PA) - SFresLines
    
    ExRespond = True

End Sub
Sub ModEnd(SheetN)
'
' подпрограмма завершения работы Модуля с листом SheetN
'  15.2.2012
'  19.4.12  - восстановление вывода Excel

    Dim i

    i = AutoFilterReset(SheetN)
    ActiveSheet.Range("A" & i).Select
    ProgressForm.Hide
' восстанавливаем вывод Excel и др
    With Application
        .StatusBar = False
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayStatusBar = True
        .DisplayAlerts = True
    End With
    ActiveSheet.DisplayPageBreaks = True
    LogWr (Doing & " - ГОТОВО!")
End Sub
Sub ErrMsg(ErrMode, MSG)
'
' - ErrMsg(ErrMode, MSG) - вывод сообщения об ощибке в Log и на экран
'                          Коды ErrMode определены в Declaration
'   31.5.12

    Dim ErrType As String, Respond As String

    Select Case ErrMode
    Case WARNING:
        LogWr "< WARNING > " & MSG
        Exit Sub
        
    Case TYPE_ERR:
        LogWr "ВНИМАНИЕ:" & MSG
        Respond = MsgBox(MSG & vbCrLf & vbCrLf & "Продолжить?", vbYesNo)
        If Respond = vbNo Then ExRespond = False
        Exit Sub
        
    Case FATAL_ERR:
Fatal:  ErrType = "<! ERROR !> "
        LogWr ErrType & MSG
        MsgBox MSG, , ErrType
'        Stop
        Exit Sub
    Case Else:
        ErrMode = FATAL_ERR
        GoTo Fatal
    End Select
End Sub
Sub LogWr(MSG)
'
' запись сообщения msg в Log-лист
'   15.2.2012

    Dim N   ' номер строки в Log
    
    N = Sheets(Log).Cells(1, 4)
    N = N + 1
    Sheets(Log).Cells(N, 1) = Date
    Sheets(Log).Cells(N, 2) = Time
    Sheets(Log).Cells(N, 3) = MSG
    Sheets(Log).Cells(1, 4) = N
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
    AutoFilterReset = Sheets(SheetN).UsedRange.Rows.count
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
Sub Pnt(Col, Criteria, Color, Optional Mode As Integer = 0)
'
' подпрограмма выбирает колонку Col по критерию Criteria и окрашивает в Color
' если Mode = 0 или не указан - окрашиваем весь ряд, иначе только Col
'   26.1.2011

    AllCol = ActiveSheet.UsedRange.Columns.count
    Range(Cells(1, 1), Cells(Lines, AllCol)).AutoFilter _
                            Field:=Col, Criteria1:=Criteria
    If Mode = 0 Then
        Range(Cells(2, 2), Cells(Lines, AllCol)).Interior.Color = Color
    Else
        Range(Cells(2, Col), Cells(Lines, Col)).Interior.Color = Color
    End If
    If Criteria = "Не состоялся" Then   ' "Не состоялся" - перечеркиваем
        Range(Cells(2, 2), Cells(Lines, AllCol)).Font.Strikethrough = True
    End If
    ActiveSheet.UsedRange.AutoFilter Field:=Col
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

    Dim s

    CurRate = 1
    If Cur = "RUB" Or Trim(Cur) = "" Then Exit Function
    s = WorksheetFunction.VLookup(Cur, Sheets(We).Range("RUB_Rate"), 2, False)
    CurRate = Replace(s, ".", ",")
End Function
Function CurISO(Cur1C)
'
' возвращает код валюты в стандарте ISO, преобразовав его из вида 1С
'   18.3.2012

    CurISO = ""
    On Error Resume Next
    CurISO = WorksheetFunction.VLookup(Cur1C, Range("Currency"), 2, False)
    On Error GoTo 0
End Function
Function DDMMYYYY(d) As String
'
' преобразование даты d в текстовый формат DDMMYYYY
'   14.2.2012
    DDMMYYYY = Day(d) & "." & Month(d) & "." & Year(d)
End Function
Function Dec(A) As String
'
' преобразование числа а в текстовый формат с десятичной точкой
'   14.2.2012

    Dec = "'" & WorksheetFunction.Substitute(A, ",", ".")
End Function
Function EOL(SheetN)
'
' возвращает количество строк в листе SheetN с учетом невидимых пустых строк
'   20/1/2012
'   4/2/2012 - обработка ошибок On Error
'   20/2/2012 - внедрение Option Explicit
'   12.5.12 - Sheets(SheetN).Select исключен

    Dim i, Col
    
    On Error GoTo Err
    
    EOL = Sheets(SheetN).UsedRange.Rows.count
    Col = Sheets(SheetN).UsedRange.Columns.count
    Do
        For i = 1 To Col
            If Sheets(SheetN).Cells(EOL, i) <> "" Then Exit Do
        Next i
        If EOL <= 1 Then Exit Do
        EOL = EOL - 1       ' иногда UsedRange оставляет пустые строки,
    Loop                    '   .. например, если в строке есть невидимый формат
    Exit Function

Err: MsgBox "Ошибка в данных на листе " & SheetN & " в ячейке (" & _
        i & "," & EOL & ")"
    Stop
End Function
Sub ClearSheet(SheetN, HDR_Range As Range)
'
' Полная очистка SheetN и перенос в него заголовка из листа Нeader.HDR_Range
'   4.2.2012
'  11.2.2012 - пересмотр спецификаций
'  10.3.12 - изменение спецификации - параметр HRD_Range
'  25.3.12 - листы C_Contr и C_ContrLnk
'  17.4.12 - лист A_Acc - новые Организации
'  18.4.12 - лист A_Dic - Словарь Организаций
'  28.4.12 - лист NewOrderList - лист Новых Заказов
'  13.5.12 - лист P_ADSKlink - новые связки Платеж - ADSK
'  15.5.12 - лист SF_PA связей Платежей с Контрактами ADSK
'   6.6.12 - Delete старый лист, создаем новый

' -- стираем старый лист
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(SheetN).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
' -- создаем новый лист
    Sheets.add After:=Sheets(Sheets.count)  ' создаем новый лист в конце справа
    ActiveSheet.Name = SheetN
    ActiveSheet.Tab.Color = RGB(50, 153, 204)   ' Tab голубой
   
    HDR_Range.Copy Sheets(SheetN).Cells(1, 1)   ' копируем шапку из Header
    
    Select Case SheetN
    Case O_NewOpp:      EOL_NewOpp = 1
    Case P_Paid:        EOL_NewPay = 1
    Case C_Contr:       EOL_NewContr = 1
    Case C_ContrLnk:    EOL_ContrLnk = 1
    Case P_PaymentUpd:  EOL_PaymentUpd = 1
    Case A_Dic:         EOL_DIC = 1
    Case A_Acc:         EOL_AdAcc = 1
    Case NewOrderList:  EOL_NewOrderList = 1
    Case P_ADSKlink:    EOL_ADSKlnkPay = 1
    Case SF_PA:         EOL_SFlnkADSK = 1
    Case NewSN:         EOL_NewSN = 1
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
        .SortFields.add key:=Cells(1, Col), SortOn:=xlSortOnValues, Order:= _
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

    Dim i, prev, x, EOL_SheetN As Integer
    
    Call SheetSort(SheetN, Col)
    EOL_SheetN = EOL(SheetN)
    
    prev = "": i = 2
    Do
        x = Sheets(SheetN).Cells(i, Col)
        If x = prev Then
            Rows(i).Delete
            EOL_SheetN = EOL_SheetN - 1
        Else
            i = i + 1
            prev = x
        End If
    Loop While i < EOL_SheetN
End Sub
Sub SheetDedup2(SheetN, ColSort, СolAcc)
'
' - SheetDedup2(SheetN, ColSort, ColAcc)  - сортируем лист SheetN
'               по колонке SortCol затем значения ColAcc сливем с "+"
'   23.5.2012

    Dim i As Integer, EOL_SheetN As Integer
    Dim prev As String, x As String
    Dim PrevAcc As String, NewAcc As String
    
    Call SheetSort(SheetN, ColSort)
    EOL_SheetN = EOL(SheetN)
    
    prev = "": i = 2
    With Sheets(SheetN)
        Do
            x = .Cells(i, ColSort)
            If x = prev Then
                PrevAcc = .Cells(i - 1, СolAcc)
                NewAcc = .Cells(i, СolAcc)
                If PrevAcc <> "" And NewAcc <> "" And PrevAcc <> NewAcc Then
                    PrevAcc = PrevAcc & "+" & NewAcc
                ElseIf PrevAcc = "" And NewAcc <> "" Then
                    PrevAcc = NewAcc
'                ElseIf PrevAcc <> "" And NewAcc = "" Then
'                ElseIf PrevAcc = "" And NewAcc = "" Then
'                   в двух последних случаях ничего не делаем
                End If
                .Cells(i - 1, СolAcc) = PrevAcc
                Rows(i).Delete
                EOL_SheetN = EOL_SheetN - 1
            Else
                i = i + 1
                prev = x
            End If
        Loop While i < EOL_SheetN
    End With
End Sub
Sub DateCol(SheetN, Col)
'
' преобразование колонки Col в листе SheetN из текста вида DD.MM.YY в формат Date
'   20.4.12

    Dim i, DD, MM, YY As Integer
    Dim Dat As Date
    Dim d() As String
    
    For i = 1 To EOL(SheetN)
        d = split(Sheets(SheetN).Cells(i, Col), ".")
        If UBound(d) = 2 Then
            DD = d(0)
            If DD < 1 Or DD > 31 Then GoTo NXT
            MM = d(1)
            If MM < 1 Or MM > 12 Then GoTo NXT
            YY = d(2)
            Dat = DD & "." & MM & "." & YY
            Sheets(SheetN).Cells(i, Col) = Dat
        End If
NXT:
    Next i
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
        
    With ProgressForm
        .ProgressFrame.Caption = Format(Pct, "0%")
        .LabelProgress.Width = Pct * .ProgressFrame.Width
        .Repaint
    End With
    
    Static T
    Dim R As String
    If T = 0 Then T = Timer
    If Timer - T > 20 Then
        T = Timer
        R = MsgBox("Дальше?", vbYesNo)
        If R = vbNo Then ExRespond = False
    End If
    
End Sub
Sub StopSub()
'
' StopSub() аварийное завершение процесса - вызывается по Событию FATAL ERROR
'
    MsgBox "Аварийная остановка StopSub", , "FATAL ERROR"
    Stop
End Sub
Function RemDelimiters(s)
'
' RemDelimeters(S) - замена спец символов в строке S на пробелы
' 19.4.12 А.Пасс

    Dim smb, i

    For i = 1 To Len(s)
        
        ' допускаем цифры, латинские или русские буквы
        smb = Mid(s, i, 1)
        If Not ((smb >= "0" And smb <= "9") _
                Or (smb >= "A" And smb <= "Z") _
                Or (smb >= "a" And smb <= "z") _
                Or (smb >= "А" And smb <= "Я") _
                Or (smb >= "а" And smb <= "я") _
                Or smb = "ъ" Or smb = "Ъ" _
                Or smb = "ё" Or smb = "Ё") Then
            Mid(s, i, 1) = " "      ' символ недопустим, заменяем на пробел
        End If
    Next i
    RemDelimiters = s
End Function
Function Compressor(s1 As Variant)
'
' удаление лишних пробелов внутри строки
'   7.3.12  из Интернет
'   7.6.12 удаляем vbCcLf

    Dim s As Variant
    s = Replace(s1, vbCrLf, " ")
    s = Trim(s)
    While InStr(1, s, "  ") <> 0
       s = Left(s, InStr(1, s, "  ") - 1) & Right(s, Len(s) - InStr(1, s, "  "))
    Wend
    Compressor = s
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

    Dim s As Range
    Dim lW As String, V As String
    
    lW = LCase$(W)
    For Each s In Range(Lst).Rows
        V = s.Cells(1, 1)
        If InStr(lW, LCase$(V)) <> 0 Then
            FindInLst = V
            Exit For
        End If
    Next s
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
    Dim s As Range
    Dim ValWord As String
    Dim Article As String

    For Each s In Range(Dic).Rows
        ValWord = s.Cells(1, 1)
        Article = s.Cells(1, 2)
        If Article = "" Then Article = ValWord
        If IsMatchList(W, Article) Then
            FindInDIC = ValWord
            Exit For
        End If
    Next s
End Function
Function IsMatchList(W, DicList) As Boolean
'
' - IsMatchList(W, DicList) - возвращает TRUE если W входит в DicList
'   7.5.12

    IsMatchList = False
    If W = "" Or DicList = "" Then Exit Function
    
    Dim x() As String
    Dim i As Integer
    Dim lW As String
    
    lW = LCase$(W)
    x = split(DicList, ",")
    
    For i = LBound(x) To UBound(x)
        If InStr(lW, LCase$(x(i))) <> 0 Then
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
