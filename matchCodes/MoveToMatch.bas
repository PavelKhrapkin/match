Attribute VB_Name = "MoveToMatch"
'---------------------------------------------------------------------------
' Ответная часть MoveToMatch, расположенная в файле match.xlsm. Загружает входные Документы.
'
' * MoveInMatch    - перенос входного Документа в базу и запуск Loader'а
'
' П.Л.Храпкин 13.09.2014

    Option Explicit    ' Force explicit variable declaration
    
Sub MoveInMatch()
Attribute MoveInMatch.VB_Description = "20.7.12 MoveToMatch of Application match2.0"
Attribute MoveInMatch.VB_ProcData.VB_Invoke_Func = "ф\n14"
'
' <*> MoveInMatch() - перемещение входного отчета в базу и запуск его обработки
'
' Получает управление по Application.Run из MoveToMatch запускаемого по Ctrl+ф
' Входной Документ (лист 1 активного файла) распознается по Штампу, соответствующего TOCmatch
'
' 18.8.12 - отделено от MoveToMatch, раположенного в PERSONAL.xlsb
' 25.8.12 - входной Документ может находиться в листе InSheetN нового загружаемого файла
' 26.8.12 - сброс окраски вместе с очисткой "1" в поле Done
' 28.8.12 - сброс Шагов, связанных с использованием загружаемого Документа
' 20.9.12 - Created Date -- исправлено для отчетов SF
' 22.12.12 - Created Date - введены переводы в ам. формат и обратно
'  6.4.13 - выход при попытке загрузить в match один из файлов базы данных
' 13.5.13 - пропускаем строки-продолжения в TOCmatch
' 17.8.13 - загрузка отчетов с частичным диапазоном дат
' 18.8.13 - ResLines теперь имеет вид 2 / 7
' 23.8.13 - SheetSort загружаемого документа, если это часть полного
' 24.8.13 - упразднил InSheetN в TOC. Теперь Документ всегда должен быть в листе 1
' 27.8.13 - минимизируем использование глобальной структуры RepTOC
' 6.10.13 - bug fix - игнорируем строки Платежа "авт нал"
' 22.10.13 - убираем избыточные .Activate; делаем Freeze Top Row
' 26.10.13 - Public RepName заменен на локальный DocName
' 16.11.13 - Обработка колонки дат и сброс фильтра дополняемого листа
' 13.09.14 - Bux fix - все еще не удален RepTOC - замена на LocalTOC
    
    Dim NewRep As String    ' имя файла с новым отчетом
    Dim DocName As String   ' имя Документа
    Dim i As Long
    Dim IsSF As Boolean     '=TRUE, если входной Документ из Salesforce
    Dim IsPartialUpdate     '=TRUE, если входной документ заменяет лишь часть отчета
    Dim FrDateTOC As Date, ToDateTOC As Date, NewFrDate As Date, NewToDate As Date
    Dim NewFrDate_Row As Long, NewFrDate_Col As Long
    Dim NewToDate_Row As Long, NewToDate_Col As Long
    Dim InSheetN As Integer 'поле в TOCmatch- номер листа входного Документа для MoveToMatch
    Dim LocalTOC As TOCmatch
    
    IsPartialUpdate = False
    NewRep = ActiveWorkbook.Name
    DocName = ActiveSheet.Name
    Lines = EOL(DocName, Workbooks(NewRep))
    
    LocalTOC = GetRep(TOC)
    
    IsSF = CheckStamp(6, NewRep, Lines)

    With DB_MATCH.Sheets(TOC)
        For i = TOCrepLines To LocalTOC.EOL
            If .Cells(i, TOC_REPNAME_COL) = "" Then GoTo NxDoc
            InSheetN = 1
''            If .Cells(i, TOC_INSHEETN) <> "" Then
''                InSheetN = .Cells(i, TOC_INSHEETN)
''            End If
            If CheckStamp(i, NewRep, Lines, IsSF, InSheetN) Then GoTo RepNameHandle
NxDoc:  Next i
    End With
    GoTo FatalInFile
        
'----- новый отчет распознан. Заменяем прежний отчет новым -----
RepNameHandle:
    Dim RepFile As String
    Dim RepLoader As String
    Dim Created As Date
    Dim MyDB As Workbook
    Dim TabColor
    
    With DB_MATCH.Sheets(TOC)
    
        If NewRep = .Cells(i, TOC_REPFILE_COL) Then
            MS "Это файл базы данных match! Его не надо загружать."
            End
        End If
        DocName = .Cells(i, TOC_REPNAME_COL)
        RepFile = .Cells(i, TOC_REPFILE_COL)
         'Lines = EOL - пятка
        Lines = Lines - GetReslines(DocName, True, .Cells(i, TOC_RESLINES_COL))
        LinesOld = .Cells(i, TOC_EOL_COL)           'EOL старого отчета
        DirDBs = .Cells(1, TOC_F_DIR_COL)
        TabColor = .Cells(i, TOC_SHEETN_COL).Interior.Color
      '--получение диапазона дат в match и новом отчете ---
        FrDateTOC = .Cells(i, TOC_FRDATE_COL)
        ToDateTOC = .Cells(i, TOC_TODATE_COL)
        NewFrDate_Row = .Cells(i, TOC_FRDATEROW_COL)
        NewFrDate_Col = .Cells(i, TOC_DATECOL_COL)
        Dim ToStr As String
        ToStr = .Cells(i, TOC_TODATEROW_COL)
        If ToStr = "EOL" Then
            NewToDate_Row = Lines
        ElseIf WorksheetFunction.IsNumber(ToStr) Then
            NewToDate_Row = ToStr
        End If
        NewToDate_Col = .Cells(i, TOC_DATECOL_COL)
        
    End With
    
    Set MyDB = Workbooks.Open(DirDBs & RepFile, UpdateLinks:=False)
    
    With Workbooks(NewRep).Sheets(InSheetN)
        If RepFile = F_SFDC Then
            Dim tst As String
            tst = .Cells(Lines + 5, 1)
            Created = GetDate(Right(.Cells(Lines + 5, 1), 16))
        ElseIf DocName = PAY_SHEET Or DocName = DOG_SHEET Then
            .Activate
            .Rows("1:" & Lines).AutoFilter
            DateCol InSheetN, NewToDate_Col
            SheetSort InSheetN, NewToDate_Col
            Dim OldDatCol As Long
            OldDatCol = PAYDATE_COL
            If DocName = DOG_SHEET Then OldDatCol = DOG1CDAT_COL
            
            DateCol DocName, OldDatCol
            SheetSort DocName, OldDatCol
            Created = GetDate(Right$(.Name, 8))
            Dim DateCell As String, PayDoc As String, Doc As Boolean
            Do
                DateCell = .Cells(NewFrDate_Row, NewFrDate_Col)
                If IsDate(DateCell) Then
                    Exit Do
                Else
                    NewFrDate_Row = NewFrDate_Row + 1
                    If NewFrDate_Row > Lines Then GoTo FatalFrDate
                End If
            Loop
            NewFrDate = GetDate(DateCell)
            Do
                Doc = True
                If DocName = PAY_SHEET Then
            '-- платежный док. !! только для Платежа
                    PayDoc = Trim(.Cells(NewToDate_Row, 1))
                    If PayDoc = "" Or InStr(PayDoc, "авт нал") <> 0 Then Doc = False
                End If
                DateCell = .Cells(NewToDate_Row, NewToDate_Col)
                If IsDate(DateCell) And Doc Then
                    Exit Do
                Else
                    NewToDate_Row = NewToDate_Row - 1
                    If NewToDate_Row < NewFrDate_Row Then GoTo FatalToDate
                End If
            Loop
            NewToDate = GetDate(DateCell)
            If NewFrDate > NewToDate Then GoTo FatalFrToDate
            If NewFrDate <> FrDateTOC Or NewToDate < ToDateTOC Then
                IsPartialUpdate = True
            End If
        ElseIf DocName = Acc1C Then
            Created = GetDate(Right$(.Cells(1, 1), 8))
        ElseIf RepFile = F_STOCK Then
            Created = GetDate(MyDB.BuiltinDocumentProperties(12))   'дата последнего Save
        Else
            Created = "0:0"
            NewFrDate = "0:0": NewToDate = "0:0"
        End If
        .UsedRange.Rows.RowHeight = 15
        .Name = "TMP"
        .Move Before:=MyDB.Sheets(DocName)
    End With
    
    With MyDB
  '-- если частичное обновление - прежний отчет не стираем, а переименовываем
  '-- .. его в *_OLD, чтобы потом слить их в Шаге MergeRep Loader'а.
  '-- .. если _OLD уже есть, но еще не обработан - уничтожаем прежний "частичный" отчет
        If IsPartialUpdate Then
            Dim OldRepName As String, sht As Worksheet
            OldRepName = DocName & "_OLD"
            If SheetExists(OldRepName) Then GoTo DelRep
            .Sheets(DocName).Name = OldRepName
        End If
DelRep: If SheetExists(DocName) Then
            Application.DisplayAlerts = False
            .Sheets(DocName).Delete
            Application.DisplayAlerts = True
        End If
        .Activate
        .Sheets("TMP").Name = DocName
        .Sheets(DocName).Tab.Color = TabColor
        .Sheets(DocName).Cells(2, 1).Select
    End With
    ActiveWindow.FreezePanes = True
    
'------------- match TOC и Log write и Save --------------
    With DB_MATCH.Sheets(TOC)
        .Cells(i, TOC_DATE_COL) = Now
        .Cells(i, TOC_EOL_COL) = Lines
        .Cells(i, TOC_MADE_COL) = REP_LOADED
        RepLoader = .Cells(i, TOC_REPLOADER_COL)
        .Cells(i, TOC_CREATED_COL) = Created
        If NewFrDate_Col > 0 Then
            .Cells(i, TOC_NEW_FRDATE_COL) = NewFrDate
            .Cells(i, TOC_NEW_TODATE_COL) = NewToDate
        End If
        .Cells(1, 1) = Now
        .Cells(1, TOC_F_DIR_COL) = DirDBs
'----------- окрашиваем даты в TOCmatch на сегодня -------------
        Dim D As Date, MaxDays As Integer
        For i = 8 To LocalTOC.EOL
            D = .Cells(i, TOC_DATE_COL)
            MaxDays = .Cells(i, TOC_MAXDAYS_COL)
            If D <> "0:00:00" And Now - D > MaxDays Then
                .Cells(i, TOC_DATE_COL).Interior.Color = vbRed
            Else
                .Cells(i, TOC_DATE_COL).Interior.Color = vbWhite
            End If
        Next i
    End With
    Dim PartStatus As String
    PartStatus = vbCrLf & "Это "
    If IsPartialUpdate Then
        PartStatus = PartStatus & "изменяемая ЧАСТЬ данных."
    Else
        PartStatus = PartStatus & "ПОЛНЫЙ документ."
    End If
    LogWr "MoveToMatch: В файл '" & RepFile & "' загружен новый отчет '" _
        & DocName & "'; EOL=" & Lines & " строк, в прежнем " & LinesOld _
        & PartStatus
        
'--- Запускаем Loader - процедуру обработки нового отчета ---
    If RepLoader <> "" Then
        ProcReset RepLoader
    End If
    MyDB.Save
    Exit Sub
    Dim Msg As String
FatalInFile:    Msg = "Не найден Штамп": GoTo FatMsg
FatalFrDate:    Msg = "FrDate": GoTo FatErMsg
FatalToDate:    Msg = "ToDate"
FatErMsg:       Msg = " не дата в ячейке " & Msg & "='" & DateCell & "'": GoTo FatMsg
FatalFrToDate:  Msg = " странные даты входного документа '" & NewRep _
                    & "': NewFrDate=" & NewFrDate & " < " & "NewToDate=" & NewToDate
FatMsg: ErrMsg FATAL_ERR, "MoveToMatch: " & Msg & vbCrLf & "Входной документ " & NewRep
End Sub
