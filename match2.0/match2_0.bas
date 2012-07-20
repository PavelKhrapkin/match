Attribute VB_Name = "match2_0"
'---------------------------------------------------------------------------
' Макросы для работы с файлом отчетов из 1С и Salesforce Match SF-1C.xlms
'
' * MoveToMatch    - перенос Листа на первое место Match1SF    (Ctrl/Shift/M)
' * TriggerOptionsFormulaStyle  - переключение моды A1/R1C1    (Ctrl/Shift/R)
'
' П.Л.Храпкин 2.1.2012
'   28.1.2012 - работы по параметризации имен и позиций листов
'    5.2.2012 - в MoveToMatch - распознавание входного отчета по штампу
'   16.5.2012 - добавлен отчет SF_PA
'    2.6.2012 - TriggerOptionsFormulaStyle A1/R1C1
'   20.7.2012 - match 2.0 - MoveToMatch с использованием TOCmatch

    Option Explicit    ' Force explicit variable declaration
    
Sub MoveToMatch()
Attribute MoveToMatch.VB_Description = "20.7.12 MoveToMatch of Application match2.0"
Attribute MoveToMatch.VB_ProcData.VB_Invoke_Func = "ф\n14"
'
' <*> MoveToMatch() - перемещение входного отчета в базу и запуск его обработки
'
' Keyboard Shortcut: Ctrl+ф     -- Ctrl/ф использован, чтобы не использовать Shift,
'                                  останавливаюший исполнение на Open
'
'Pavel Khrapkin 23-Dec-2011
' 8.2.2012 - распознаем новый отчет, запускаем его обработку
' 20.7.12 - match2.0 - распознавание отчета, перенос его в один из файлов базы и запуск обработки

    Dim NewRep As String            ' имя файла с новым отчетом
    Dim i As Integer
    
    NewRep = ActiveWorkbook.Name
    Lines = EOL(1, Workbooks(NewRep))
    
    Set DB_MATCH = Workbooks.Open(F_MATCH, UpdateLinks:=False)
    
'------ распознавание Штампа файла NewRep по таблице TOCmatch -------------
                'в (4, TOC_HANDLE_COL) TOCmatch - число баз данных
                '.. в одну из которых перепишем новый отчет
    For i = 5 To 5 + DB_MATCH.Sheets(TOC).Cells(4, TOC_CREATED_COL)
        If IsThisStamp(i, NewRep) Then Exit For
    Next i
  
'------ распознавание RepName по таблице TOCmatch -------------
    Dim FrTOC As Integer, ToTOC As Integer  'строки поиска RepName в TOC
    With DB_MATCH.Sheets(TOC)
        FrTOC = .Cells(i, TOC_FRTOC_COL)
        ToTOC = .Cells(i, TOC_TOTOC_COL) + FrTOC - 1
    End With
    For i = FrTOC To ToTOC
        If IsThisStamp(i, NewRep) Then GoTo RepNameHandle
    Next i
    GoTo FatalNewRep
        
'----- новый отчет распознаван. Заменяем прежний отчет новым -----
RepNameHandle:
    Dim RepFile As String
    Dim RepLoader As String
    Dim Created As Date
    Dim MyDB As Workbook
    
    With DB_MATCH.Sheets(TOC)
        Lines = Lines - .Cells(i, TOC_RESLINES_COL) '= EOL - пятка
        LinesOld = .Cells(i, TOC_EOL_COL)           'EOL старого отчета
        RepFile = .Cells(i, TOC_REPDIR_COL) & .Cells(i, TOC_REPFILE_COL)
        RepName = .Cells(i, TOC_REPNAME_COL)
    End With
    
'    Set MyDB = Workbooks.Open(RepFile, UpdateLinks:=False)
    Set MyDB = Workbooks.Open(RepFile)
    
    With Workbooks(NewRep).Sheets(1)
        If RepFile = F_SFDC Then
            Created = Mid(.Cells(Lines + 5, 1), 24)
        Else
            Created = "1.1.1900"
        End If
        .UsedRange.Rows.RowHeight = 15
        .Name = "TMP"
        .Move Before:=MyDB.Sheets(RepName)
    End With
    
    With MyDB
        .Activate
        Application.DisplayAlerts = False
        .Sheets(RepName).Delete
        Application.DisplayAlerts = True
        .Sheets("TMP").Name = RepName
        .Sheets(RepName).Tab.Color = rgbViolet
    End With
    
    LogWr "MoveToMatch: Загружен новый отчет " & RepName _
        & "; EOL=" & Lines & " строк, в прежнем " & LinesOld
'------------- match TOC и Log write и Save --------------
    With DB_MATCH.Sheets(TOC)
        .Activate
        .Cells(i, TOC_DATE_COL) = Now
        .Cells(i, TOC_CREATED_COL) = ""
        .Cells(i, TOC_EOL_COL) = Lines
        .Cells(i, TOC_MADE_COL) = REP_LOADED
        RepLoader = .Cells(i, TOC_REPLOADER_COL)
        .Cells(i, TOC_CREATED_COL) = Created
        .Cells(i, TOC_NEXTREP_COL) = ""
        .Cells(1, 1) = Now
    End With
    LogWr "Новый отчет '" & RepName & "' загружен в " & RepFile
    DB_MATCH.Save
    DB_MATCH.Close
'--- Запускаем Loader - процедуру обработки нового отчета ---
    If RepLoader <> "" Then
        Application.Run ("'" & RepFile & "'!" & RepLoader)
    End If
    MyDB.Save
    Close
    End
FatalNewRep:
    ErrMsg FATAL_ERR, "Входной отчет '" & NewRep & "' не распознан"
End Sub
Function IsThisStamp(iTOC, NewRep) As Boolean
'
' - IsThisStamp(iTOC) - проверка соответствия нового отчета штампу в строке iTOC.
' 19.7.2012

    Dim NewRepStamp As String       ' штамп нового отчета
    
    Dim Stamp As String         '= строка - штамп
    Dim StampType As String     'тип штампа: строка (=) или подстрока
    Dim Stamp_R As Integer      'номер строки, где штамп
    Dim Stamp_C As Integer      'номер колонки, где штамп
    Dim ParCheck As Integer     'параметр TOCmatch - строка дополнительной проверки штампа
    
    IsThisStamp = False
    RepName = ""
        
    With DB_MATCH.Sheets(TOC)
        Do
            Stamp = .Cells(iTOC, TOC_STAMP_COL)
            If Stamp = "" Then Exit Function        ' отсутствует штамп - не годится!
            StampType = .Cells(iTOC, TOC_STAMP_TYPE_COL)
            Stamp_R = .Cells(iTOC, TOC_STAMP_R_COL)
            Stamp_R = Stamp_R + Lines - .Cells(iTOC, TOC_RESLINES_COL)
            Stamp_C = .Cells(iTOC, TOC_STAMP_C_COL)
            NewRepStamp = Workbooks(NewRep).Sheets(1).Cells(Stamp_R, Stamp_C)
            
            If StampType = "=" And NewRepStamp <> Stamp Then
                Exit Function
            ElseIf StampType = "I" And InStr(LCase$(NewRepStamp), LCase$(Stamp)) = 0 Then
                Exit Function
            Else: If StampType <> "=" And StampType <> "I" Then _
                ErrMsg FATAL_ERR, "Сбой в структоре TOCmatch: тип штампа =" & StampType
            End If
        
            ParCheck = .Cells(iTOC, TOC_PARCHECK_COL)
            If IsNumeric(ParCheck) And ParCheck > 0 Then iTOC = ParCheck
        Loop While ParCheck <> 0
        RepName = .Cells(1, TOC_REPNAME_COL)
    End With

    IsThisStamp = True

End Function
Sub TriggerOptionsFormulaStyle()
'
' * Trigger Options-Formula Style A1/R1C1
'
' CTRL+Shift+R
'
' 2.6.12
    If Application.ReferenceStyle = xlR1C1 Then
        Application.ReferenceStyle = xlA1
    Else
        Application.ReferenceStyle = xlR1C1
    End If
End Sub
