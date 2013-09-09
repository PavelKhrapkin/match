Attribute VB_Name = "TOCengine"
'---------------------------------------------------------------------------------------
' TOCengine - процессор TOC - Table Of Content Документов в match.xlsx
'
' 28.08.2013
'=========================== Описания ================================
'       * TOC храниться в листе TOC. Данные о Документе хранятся в виде строки этого листа
'       * при работе отдельных Шагов, параметры и константы передаются в структуре TOCmatch,
'         заполняемой данными при открытии Документа процедурой GetRep.
'       * во время извлечения Документа делается проверка его Штампов (CheckStamp)
'       * часть данных TOC - констант - не изменяется, переменные же записываются WrTOC
'       * процедура StepIn вызывает все нужные для работы Шага Документы, в при завершении
'         Шага автоматически вызывается StepOut, выполняющий WrTOC
'       * Такая политика делает отдельные Шаги контектстно независимыми: глабальные данные
'         не передаются между Документами и Шагами, помимо интерфейса модуля TOCengine
'
'============================ Процедуры =============================
' - GetRep(RepName)             - находит и проверяет штамп отчета RepName
' - FatalRep(SubName, RepName)  - сообщение о фатальной ошибке при запросе RepName
' - CheckStamp(iTOC, [FromMoveToMatch]) - проверка Штампа по стоке в TOCmatch
' - FileOpen(RepFile)           - проверяет, открыт ли RepFile, если нет - открывает
' - GetReslines(x,LoadMode)     - извлечение размера пятки из х с учетом контекста LoadMode
' - WrTOC([Name])               - записывает Public RepTOC в TOCmatch для Документа Name

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
'   13.8.13 - добавлено поле iTOC в структуру TOCmatch - номер строки в TOC
'   18.8.13 - с подпрограммой GetReslines - изменение размера пятки при загрузке и далее
'   27.8.13 - не используем глобальную структуру RepTOC

    Const TOClineN = 4  ' номер строки в TOCmatch описывающей саму себя
    Dim i As Long, EOL_TOC As Long, RepFile As String
    
    If RepName = "" Then Exit Function
    
    If DB_MATCH Is Nothing Then
        Set DB_MATCH = FileOpen(F_MATCH)
        EOL_TOC = EOL(TOC, DB_MATCH)
        DB_MATCH.Sheets(TOC).Cells(TOClineN, TOC_EOL_COL) = EOL_TOC
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
        For i = TOClineN To EOL_TOC
            If .Cells(i, TOC_REPNAME_COL) = RepName Then GoTo FoundRep
        Next i
        FatalRep "GetRep ", RepName

FoundRep:
        RepFile = .Cells(i, TOC_REPFILE_COL)
    End With
    
'---- проверка штампа ----------
    Dim Str As Long, StC As Long
    Dim TestedStamp As String
    
    Select Case RepFile
    Case F_MATCH:
    Case F_1C:      Set DB_1C = FileOpen(RepFile)
    Case F_SFDC:    Set DB_SFDC = FileOpen(RepFile)
    Case F_ADSK:    Set DB_ADSK = FileOpen(RepFile)
    Case F_STOCK:   Set DB_STOCK = FileOpen(RepFile)
    Case F_TMP:     Set DB_TMP = FileOpen(RepFile)
    Case Else: FatalRep "GetRep: файл штампа=" & RepFile, RepName
    End Select
            
    If CheckStamp(i) Then
        GetRep = GetTOC(i)
    Else
        FatalRep "GetRep", RepName
    End If
End Function
Function GetTOC(ByVal iTOC As Long) As TOCmatch
'
'  - GetTOC(iTOC)   - возвращает заполненную структуру TOCmatch по строке iTOC в TOC
'
' 27.8.13

    Dim LocalTOC As TOCmatch
    Dim LoadMode As Boolean
    LoadMode = False
    
    With DB_MATCH.Sheets(TOC)
        LocalTOC.iTOC = iTOC                ' номер строки в TOC - Read Only!
        LocalTOC.Dat = .Cells(iTOC, TOC_DATE_COL)
        LocalTOC.Name = .Cells(iTOC, TOC_REPNAME_COL)
        LocalTOC.MyCol = .Cells(iTOC, TOC_MYCOL_COL)
        If LocalTOC.Made = REP_LOADED Then LoadMode = True
        LocalTOC.ResLines = GetReslines(, LoadMode, .Cells(iTOC, TOC_RESLINES_COL))
        LocalTOC.Made = .Cells(iTOC, TOC_MADE_COL)
        LocalTOC.RepFile = .Cells(iTOC, TOC_REPFILE_COL)
        LocalTOC.SheetN = .Cells(iTOC, TOC_SHEETN_COL)
        LocalTOC.EOL = .Cells(iTOC, TOC_EOL_COL)
        LocalTOC.CreateDat = .Cells(iTOC, TOC_CREATED_COL)
        LocalTOC.FormName = .Cells(iTOC, TOC_FORMNAME)
    End With
    GetTOC = LocalTOC
End Function

Sub FatalRep(SubName, RepName)
'
' - FatalRep(SubName, RepName) - сообщение о фатальной ошибке при запросе RepName
' 17.8.12
' 9.8.12 -- более ясная диагностика по не найденному Штампу

    ErrMsg FATAL_ERR, SubName & "> Не найден Штамп в Документе '" & RepName & "'" _
        & vbCrLf & vbCrLf & "Этот Документ надо загрузить в match заново или " _
        & vbCrLf & "исправить параметры в TOCmatch."
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
'  6.4.13 - обработка Exception при поиске Штампа. Ошибка - значит Штампа нет.
' 14.7.13 - дополнительная диагностика и действия, если Штамп не найден
' 27.8.13 - минимизируем использование глобальной структуры RepTOC за счет GetTOC(iTOC)

    Dim SR() As String, SC() As String
    Dim Str As Long, StC As Long
    
    Dim RepName As String
    Dim txt As String, TestedStamp As String
    Dim Typ As String
    Dim Continued As String
    Dim i As Long, j As Long
    Dim LocalTOC As TOCmatch
    
    LocalTOC = GetTOC(iTOC)
    
    On Error GoTo NoStamp
    CheckStamp = True
    
    With DB_MATCH.Sheets(TOC)
        SR = split(.Cells(iTOC, TOC_STAMP_R_COL), ",")
        SC = split(.Cells(iTOC, TOC_STAMP_C_COL), ",")
        txt = .Cells(iTOC, TOC_STAMP_COL)
        Typ = .Cells(iTOC, TOC_STAMP_TYPE_COL)
        If Typ = "N" Then GoTo ex
        RepName = .Cells(iTOC, TOC_REPNAME_COL)
        Continued = .Cells(iTOC, TOC_PARCHECK_COL)
    End With
    
    With LocalTOC
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
ex:             Exit Function
NxtChk:
            Next j
        Next i
        If NewRep = "" Then
            Dim ToChangeEOLinTOC As String, RightEOL As Long
            RightEOL = EOL(.SheetN) - .ResLines
            ToChangeEOLinTOC = MsgBox("CheckStamp: не нашли Штамп '" & txt & "' в строке " & Str _
                & vbCrLf & "полагая,  что EOL = " & .EOL & ";" _
                & vbCrLf & "на самом деле EOL = " & RightEOL _
                & vbCrLf & vbCrLf & "Исправить EOL в TOCmatch? ", vbYesNo)
            If ToChangeEOLinTOC = vbYes Then
                LocalTOC.EOL = RightEOL
                RepTOC = LocalTOC
                WrTOC
                CheckStamp (iTOC)
                Exit Function
            Else
                FatalRep "GetRep.CheckStamp", RepName
            End If
        End If
NoStamp: CheckStamp = False
    End With
End Function
Function FileOpen(RepFile) As Workbook
'
' - FileOpen(RepFile)   - проверяет, открыт ли RepFile, если нет - открывает
'   26.7.12
'   28.8.13 - DisplayAlerts False чтобы не сообщать о пере-открытии файла
    
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
    
    Application.DisplayAlerts = False
    Set FileOpen = Workbooks.Open(DirDBs & RepFile, UpdateLinks:=False)
    Application.DisplayAlerts = True
End Function
Function GetReslines(Optional ByVal Doc As String, _
    Optional ByVal LoadMode As Boolean = False, Optional Resl As String = "?") As Long
'
' - GetReslines([Doc],[LoadMode],[ResL]) - извлечение размера пятки Doc с учетом
'            контекста LoadMode; строка ResL со значениями размера пятки
'            может быть явно указана в обращении, чтобы ее не искать повторно
'
' ! таким образом первый Шаг после загрузки документа должен добавлять пятку в ResLines
'
' 18.8.13
' 19.8.13 - по умолчанию LoadMode проверяем статус документа в RepTOC.Made

        Dim ss() As String, R As TOCmatch
        
        GetReslines = 0
        If Resl = "" Then Exit Function
        If Resl = "?" Then
            If Doc = "" Or Doc = TOC Or Doc = Process Then Exit Function
            If IsMissing(Doc) Then FatalRep "GetResLines", Doc
            R = GetRep(Doc)
            Resl = DB_MATCH.Sheets(TOC).Cells(R.iTOC, TOC_RESLINES_COL)
        End If
'''        If Resl = "" Then Exit Function
        
        If InStr(Resl, "/") <> 0 Then
            ss = split(Resl, "/")
            If IsMissing(LoadMode) Then
                LoadMode = False
                If R.Made = REP_LOADED Then LoadMode = True
            End If
            If LoadMode Then
                GetReslines = ss(0)
            Else
                GetReslines = ss(UBound(ss))
            End If
        ElseIf IsNumeric(Resl) Then
            GetReslines = Resl
        End If
End Function

Sub WrTOC(Optional ByVal Name As String = "")
'
' - WrTOC([Name]) - записывает данные по документу Name в оглавление
'                   По умолчанию Name последнего открытого GetRep документа
'
'     * записываются не все данные из WrTOC. Некоторые элементы структуры, например,
'       данные Штампа, являются Read Only
'
'   5.8.2012
'  12.8.12 - "серые" колонки описывающие Штамп не записываем
'  17.8.12 - еще ряд полей не записывыем в match.xlsm и использование FatalRep
'   2.9.12 - дополнительные ограничения записи в TOCmatch
' 28.10.12 - записывает в TOCmatch дату создания CreateDat
' 14.07.13 - Save Changes в DBs
' 15.08.13 - Optional Name - имя документа, по которому сохраняем строку TOCmatch
' 26.08.13 - переписываем EOL в TOC
' 28.08.13 - Name из RepTOC.Name

    Dim i As Long
    Const BEGIN = 8 ' начало списка обрабатываемых Документов
    
    If Name = "" Then Name = RepTOC.Name    ' по умолчанию Name по последнему GetRep
    If Name = "" Then FatalRep "WrTOC", "<пусто>"
    If Name = TOC Then Exit Sub
    
    For i = BEGIN To BIG
        If DB_MATCH.Sheets(1).Cells(i, TOC_REPNAME_COL) = Name Then GoTo FoundRep
    Next i
    FatalRep "WrTOC", Name

FoundRep:
    Dim NewEOL As Long, W As String, S As String
    W = DB_MATCH.Sheets(TOC).Cells(i, TOC_REPFILE_COL)
    S = DB_MATCH.Sheets(TOC).Cells(i, TOC_SHEETN_COL)
    Call FileOpen(W)
    Workbooks(W).Sheets(S).Activate
    Dim Resl As String
    Resl = DB_MATCH.Sheets(1).Cells(i, TOC_RESLINES_COL)
    NewEOL = EOL(Name) - GetReslines(Name, , Resl)
    If NewEOL <= 0 Then GoTo Err
    With DB_MATCH.Sheets(TOC)
'        .Cells(i, TOC_EOL_COL) = EOL(Name, Workbooks(RepTOC.RepFile)) - RepTOC.ResLines
        .Cells(i, TOC_EOL_COL) = NewEOL
        If Not CheckStamp(i) Then GoTo Err
        .Cells(i, TOC_DATE_COL) = RepTOC.Dat
        .Cells(i, TOC_MADE_COL) = RepTOC.Made
'''        .Cells(i, TOC_MYCOL_COL) = RepTOC.MyCol
'''        .Cells(i, TOC_RESLINES_COL) = RepTOC.ResLines
'''        .Cells(i, TOC_REPFILE_COL) = RepTOC.RepFile
'''        .Cells(i, TOC_SHEETN_COL) = RepTOC.SheetN
'''        .Cells(i, TOC_STAMP_COL) = RepTOC.Stamp
'''        .Cells(i, TOC_STAMP_TYPE_COL) = RepTOC.StampType
'''        .Cells(i, TOC_STAMP_R_COL) = RepTOC.StampR
'''        .Cells(i, TOC_STAMP_C_COL) = RepTOC.StampC
'''        .Cells(i, TOC_CREATED_COL) = RepTOC.CreateDat
'''        .Cells(i, TOC_PARCHECK_COL) = RepTOC.ParChech
'''        .Cells(i, TOC_REPLOADER_COL) = RepTOC.Loader
        .Cells(1, 1) = Now
    End With
    DB_MATCH.Save
    Exit Sub
Err: ErrMsg FATAL_ERR, "WrTOC> На Шаге '" & RepTOC.Made _
        & "' Документ '" & Name & "' странно изменил EOL=" & RepTOC.EOL
End Sub

