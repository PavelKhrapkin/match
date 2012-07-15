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

    Option Explicit    ' Force explicit variable declaration
    
Sub MoveToMatch()
Attribute MoveToMatch.VB_Description = "8.2.2012 - перемещение входного отчета на первый лист MatchSF-1C.xlsb,  распознавание его по штампу и запуск макроса по его замене "
Attribute MoveToMatch.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' <*> MoveToMatch() - перемещение входного отчета в базу и запуск его обработки
'
' Keyboard Shortcut: Ctrl+Shift+M
'
'Pavel Khrapkin 23-Dec-2011
' 8.2.2012 - распознаем новый отчет, запускаем его обработку
' 11.7.12 - match2.0 - распознавание отчета, перенос его в один из файлов базы и запуск обработки

    Dim NewRep As String            ' имя файла с новым отчетом
    Dim i As Integer
    
    NewRep = ActiveWorkbook.Name
    Lines = EOL(1, Workbooks(NewRep))

    Set DB_MATCH = Workbooks.Open(F_MATCH, UpdateLinks:=False)
    
'------ распознавание Штампа файла NewRep по таблице TOCmatch -------------
    Dim TOCline As Range        '= строка TOC match
                    'в строке 4 TOCmatch - TOCmatch количество баз данных
    For i = 5 To 5 + DB_MATCH.Sheets(TOC).Cells(4, TOC_PAR_1_COL)
        Set TOCline = DB_MATCH.Sheets(TOC).Rows(i)
        If IsRightStamp(TOCline, NewRep, 1, True) Then Exit For
    Next i
  
'------ распознавание RepName по таблице TOCmatch -------------
    With TOCline
        Dim FrTOC As Integer, ToTOC As Integer  'строки поиска RepName в TOC
        FrTOC = .Cells(1, TOC_FRTOC_COL)
        ToTOC = .Cells(1, TOC_TOTOC_COL) + FrTOC - 1
        For i = FrTOC To ToTOC
            Set TOCline = DB_MATCH.Sheets(TOC).Rows(i)
            If IsRightStamp(TOCline, NewRep, 1, True) Then GoTo RepNameHandle
        Next i
        GoTo FatalNewRep
        
'----- новый отчет распознаван. Заменяем прежний отчет новым -----
RepNameHandle:
        Dim RepFile As String
        Dim RepLoader As String
        Dim MyDB As Workbook
                
        Lines = Lines - Val(.Cells(1, TOC_RESLINES_COL))  '= EOL - пятка
        RepFile = .Cells(1, TOC_REPDIR_COL) & .Cells(1, TOC_REPFILE_COL)
        Set MyDB = Workbooks.Open(RepFile, UpdateLinks:=False)
        
        With Workbooks(NewRep).Sheets(1)
            .UsedRange.Rows.RowHeight = 15
            .Name = "TMP"
            .Move Before:=MyDB.Sheets(RepName)
        End With
        With MyDB
            .Activate
            .Sheets(RepName).Delete
            .Sheets("TMP").Name = RepName
            .Sheets(RepName).Tab.Color = rgbViolet
        End With
        
'--- Запускаем Loader - процедуру обработки нового отчета ---
        RepLoader = TOCline.Cells(1, TOC_REPLOADER_COL)
        If RepLoader <> "" Then
            Application.Run ("'" & RepFile & "'!" & RepLoader)
        End If
    End With
    LogWr "Загружен новый отчет " & RepName
    MyDB.Save
    MyDB.Close
'------------- match TOC и Log write и Save --------------
    With DB_MATCH.Sheets(TOC)
        .Activate
        .Cells(i, TOC_DATE_COL) = Now
        .Cells(i, TOC_HANDLE_COL) = ""
        .Cells(i, TOC_EOL_COL) = Lines
        .Cells(1, 1) = Now
    End With
    LogWr "Новый отчет '" & RepName & "' загружен в " & RepFile
    DB_MATCH.Save
    DB_MATCH.Close
    Exit Sub
FatalNewRep:
    ErrMsg FATAL_ERR, "Входной отчет '" & NewRep & "' не распознан"
End Sub
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
