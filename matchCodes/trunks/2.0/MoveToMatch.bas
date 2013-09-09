Attribute VB_Name = "MoveToMatch"
'---------------------------------------------------------------------------
' Ответная часть MoveToMatch, расположенная в файле match.xlsm. Загружает входные Документы.
'
' * MoveInMatch    - перенос входного Документа в базу и запуск Loader'а
'
' П.Л.Храпкин 27.8.2013

    Option Explicit    ' Force explicit variable declaration
    
Sub MoveInMatch()
Attribute MoveInMatch.VB_Description = "20.7.12 MoveToMatch of Application match2.0"
Attribute MoveInMatch.VB_ProcData.VB_Invoke_Func = "ф\n14"
'
' <*> MoveToMatch() - перемещение входного отчета в базу и запуск его обработки
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
    
    Dim NewRep As String    ' имя файла с новым отчетом
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
    RepName = ActiveSheet.Name
    Lines = EOL(RepName, Workbooks(NewRep))
    
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
        RepName = .Cells(i, TOC_REPNAME_COL)
        RepFile = .Cells(i, TOC_REPFILE_COL)
         'Lines = EOL - пятка
        Lines = Lines - GetReslines(RepName, True, .Cells(i, TOC_RESLINES_COL))
        LinesOld = .Cells(i, TOC_EOL_COL)           'EOL старого отчета
        DirDBs = .Cells(1, TOC_F_DIR_COL)
        TabColor = .Cells(i, TOC_SHEETN_COL).Interior.Color
      '--получение диапазона дат в match и новом отчете ---
        FrDateTOC = .Cells(i, TOC_FRDATE_COL)   ' Даты прежнего отчета
        ToDateTOC = .Cells(i, TOC_TODATE_COL)   '.. в Match
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
        ElseIf RepName = PAY_SHEET Or RepName = DOG_SHEET Then
            .Activate
            .Rows("1:" & Lines).AutoFilter
            DateCol InSheetN, NewToDate_Col
            SheetSort InSheetN, NewToDate_Col
            Created = GetDate(Right$(.Name, 8))
            Dim DateCell As String
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
                DateCell = .Cells(NewToDate_Row, NewToDate_Col)
                If IsDate(DateCell) Then
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
        ElseIf RepName = Acc1C Then
            Created = GetDate(Right$(.Cells(1, 1), 8))
        ElseIf RepFile = F_STOCK Then
            Created = GetDate(MyDB.BuiltinDocumentProperties(12))   'дата последнего Save
        Else
            Created = "0:0"
            NewFrDate = "0:0": NewToDate = "0:0"
        End If
        .UsedRange.Rows.RowHeight = 15
        .Name = "TMP"
        .Move Before:=MyDB.Sheets(RepName)
    End With
    
    With MyDB
        .Activate
  '-- если частичное обновление - прежний отчет не стираем, а переименовываем
  '-- .. его в *_OLD, чтобы потом слить их в Шаге MergeRep Loader'а.
  '-- .. если _OLD уже есть, но еще не обработан - уничтожаем прежний "частичный" отчет
        If IsPartialUpdate Then
            Dim OldRepName As String, sht As Worksheet
            OldRepName = RepName & "_OLD"
            If SheetExists(OldRepName) Then GoTo DelRep
            .Sheets(RepName).Name = OldRepName
        End If
DelRep: If SheetExists(RepName) Then
            Application.DisplayAlerts = False
            .Sheets(RepName).Delete
            Application.DisplayAlerts = True
        End If
        .Sheets("TMP").Name = RepName
        .Sheets(RepName).Tab.Color = TabColor
    End With
    
'------------- match TOC и Log write и Save --------------
    With DB_MATCH.Sheets(TOC)
        .Activate
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
        For i = 4 To RepTOC.EOL
            D = .Cells(i, TOC_DATE_COL)
            MaxDays = .Cells(i, TOC_MAXDAYS_COL)
            If D <> "0:00:00" And Now - D > MaxDays Then
                .Cells(i, TOC_DATE_COL).Interior.Color = vbRed
            Else
                .Cells(i, TOC_DATE_COL).Interior.Color = vbWhite
            End If
        Next i
    End With
'---------- Сброс всех Процессов, работающих с загружаемым Документом
    With DB_MATCH.Sheets(Process)
        .Activate
        For i = 6 To EOL(Process, DB_MATCH)
            If .Cells(i, PROC_REP1_COL) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 1) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 2) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 3) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 4) = RepName Then
                StepReset i
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
        & RepName & "'; EOL=" & Lines & " строк, в прежнем " & LinesOld _
        & PartStatus
        
'--- Запускаем Loader - процедуру обработки нового отчета ---
    If RepLoader <> "" Then
        ProcStart RepLoader
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
Sub StepReset(iStep)
'
' - StepReset(iStep) - сброс Шага в таблице Процессов - РЕКУРСИЯ!
' 28.8.12
'  9.9.12 - bug fix в сбосе выполненного Шага при загрузке нового Документа
' 13.9.12 - bug fix - не сбрасываем Шаги <*>ProcStart

    Dim Step As String, PrevStep As String
    Dim Proc As String, ThisProc As String
    Dim i As Integer, iProc As Integer
    
    With DB_MATCH.Sheets(Process)
        If .Cells(iStep, PROC_STEPDONE_COL) = "" Then Exit Sub
        Step = .Cells(iStep, PROC_STEP_COL)
'---- сброс Шага iStep и окраски старта его Процедуры "<*>ProcStart"
        For i = 6 To EOL(Process, DB_MATCH)
            If .Cells(i, PROC_STEP_COL) = PROC_START Then iProc = i
            If i = iStep Then
                If .Cells(i, PROC_STEPDONE_COL) = "1" Then ' пропускаем <*>ProcStart
                    .Cells(i, PROC_STEPDONE_COL) = ""
                End If
                .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 0
                .Range(Cells(iProc, 1), Cells(iProc, 3)).Interior.ColorIndex = 0
                Exit For
            End If
        Next i
'---- сброс окраски конца Процедуры "<*>ProcEnd"
        For i = iProc + 1 To EOL(Process, DB_MATCH)
            If .Cells(i, PROC_STEP_COL) = PROC_END Then
                .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 0
                Exit For
            End If
        Next i
'---- сброс Шагов, в которых в PrevStep ссылаются на Шаг в "своем" Процессе
        Proc = .Cells(iProc, PROC_NAME_COL)
        For i = iProc + 1 To EOL(Process, DB_MATCH)
            PrevStep = .Cells(i, PROC_PREVSTEP_COL)
            If InStr(PrevStep, Step) <> 0 And i <> iStep Then
                StepReset i                                     '* РЕКУРСИЯ *
            End If
            If .Cells(i, PROC_STEP_COL) = PROC_END Then Exit For
        Next i
'---- сброс Шагов, в которых в PrevStep ссылаются на Шаг из "другого" Процесса
        For i = 2 To EOL(Process, DB_MATCH)
            PrevStep = .Cells(iStep, PROC_PREVSTEP_COL)
            ThisProc = .Cells(iStep, PROC_NAME_COL)
            If InStr(PrevStep, Proc & "/" & Step) Then StepReset i '* РЕКУРСИЯ *
        Next i
    End With
End Sub
