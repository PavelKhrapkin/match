Attribute VB_Name = "For_PERSONAL"
'---------------------------------------------------------------------------
' Макросы для работы с файлом отчетов из 1С и Salesforce Match SF-1C.xlms
'
' * MoveToMatch    - перенос Листа на первое место Match1SF    (Ctrl/Shift/M)
' * TriggerOptionsFormulaStyle  - переключение моды A1/R1C1    (Ctrl/Shift/R)
'
' П.Л.Храпкин 21.8.2014
'   28.1.2012 - работы по параметризации имен и позиций листов
'    5.2.2012 - в MoveToMatch - распознавание входного отчета по штампу
'   16.5.2012 - добавлен отчет SF_PA
'    2.6.2012 - TriggerOptionsFormulaStyle A1/R1C1
'   26.7.2012 - match 2.0 - MoveToMatch с использованием TOCmatch
'   17.8.2012 - Обработка Процессов - Loader'ов в ProcessEngine
'    8.9.2012 - этот модуль помещен под названием ForPERSONAL.bas, чтобы не путать
'   9.11.2013 - добавлены макросы для обработки отчетов SN из PartnerCenter.Autodesk
'   21.8.2014 - используем Dir(P) в MoveToMatch

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
' 26.7.12 - match2.0 - распознавание отчета по ТОС
' 1.8.12 - RepTOC.EOL вместо вызова EOL(TOC,DB_MATCH), bug fix
'          Сброс всех Процессов, работающих с загружаемым Документом
' 11.8.12 - bug fix - раскраска даты отчета
' 18.8.12 - перенос основного кода в MoveInMatch в match.xlsm
' 10.9.12 - bug fix - не там брал Path DBs
' 21.8.14 - проверка, что match.xlsm есть по Dir(P)
    
    Dim D As String
    D = "C:\work\Match\match2.0\DBs"
    Const F = "match.xlsm"
    Const DinC = "C:\match_environment.xlsx"
    Dim P As String
    
    Dim NewRep As String    ' имя файла с новым отчетом
    NewRep = ActiveWorkbook.Name
    If NewRep = "Book1" Or NewRep = "PERSONAL.XLSB" Then
        MsgBox "Открыт пустой входной документ или PERSONAL.xlsb", , "<ERROR!>"
        End
    End If

OpenTry:
    P = D & "\" & F
    
'------- файл match.xlsm есть месте записанном в D?
    Dim W As Workbook
    On Error Resume Next
    Set W = Workbooks.Open(P, UpdateLinks:=False)
    On Error GoTo 0

    If W Is Nothing Or Dir(P) = "" Then
    '------- нет, match.xlsm не там. Он есть среди Workbooks?
        For Each W In Application.Workbooks
            If W.Name = F Then
                P = W.Path & "\" & F
                GoTo RunMatch
            End If
        Next W
'------ и среди Workbooks нет. Посмотрим в файле в С:\
        If Dir(DinC) = "" Then
            MsgBox "Не доступен файл '" & DinC & "'!"
            End
        End If
        On Error Resume Next
        Set W = Workbooks.Open(DinC)
        P = W.Sheets(1).Cells(1, 2) & F
        W.Close
        Set W = Workbooks.Open(P, UpdateLinks:=False)
        On Error GoTo 0
        If Dir(P) = "" Then
Const Msg = "<!> MoveToMatch не удалось открыть файл match.xlsm'" _
    & vbCrLf & vbCrLf & "Попробуй открыть его вручную, а потом" _
    & vbCrLf & "еще раз запусти загрузку Документа"
            If MsgBox(Msg, vbYesNo) = vbYes Then GoTo OpenTry
            End
        End If
    End If
RunMatch:
    Workbooks(NewRep).Activate
    Application.Run "'" & P & "'!MoveInMatch"
        
    End Sub
Sub TriggerOptionsFormulaStyle()
Attribute TriggerOptionsFormulaStyle.VB_ProcData.VB_Invoke_Func = "R\n14"
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
Sub AdskSN()
'
' AdskSN Macro
' Обновление еженедельного отчета по SN из PartnerCenter.Autodesk.com
'
' 9.11.13 П.Храпкин

    Dim PrevRep As Workbook, SFrep As Range
    Dim NewRepName As String
    Dim EOLrep As Long, EOLprev As Long
        
    With ActiveWorkbook
        
        .Sheets.Add After:=.Sheets(1)
        ActiveSheet.Name = "SF"
        Set PrevRep = GetPrevRep()
        Set SFrep = PrevRep.Sheets("SF").UsedRange
        With .Sheets("SF")
            .Tab.Color = vbMagenta
            SFrep.Copy .Cells(1, 1)
            .Activate
            Rows("2:2257").Select
            Selection.RowHeight = 15
            Columns("A:A").ColumnWidth = 12
            Columns("H:H").ColumnWidth = 11
            Columns("I:I").ColumnWidth = 20
            Columns("K:K").ColumnWidth = 11
            Columns("L:L").ColumnWidth = 11
        End With
        With .Sheets(1)
            .Activate
            EOLrep = .UsedRange.Rows.Count
            Do While .Cells(EOLrep, 6) = ""
                EOLrep = EOLrep - 1
            Loop
            EOLprev = PrevRep.Sheets(1).UsedRange.Rows.Count
            .Columns("A:A").Insert
            .Columns("A:A").Insert
            .Columns("A:A").Insert
            .Columns("A:A").Insert
            PrevRep.Sheets(1).Columns("A:D").Copy .Cells(1, 1)
            .Columns("A:D").ColumnWidth = 4
            .Cells(6, 4).FormulaR1C1 = "=IF(ISERROR(VLOOKUP(RC[1],SF!C,1,FALSE)),"""",1)"
            .Cells(6, 3).FormulaR1C1 = "=IF(ISERROR(VLOOKUP(RC[18],SF!C[5],1,FALSE)),"""",1)"
            .Cells(6, 2).FormulaR1C1 = _
                "=IF(ISERROR(VLOOKUP(RC[51],SF!C[17],1,FALSE)),"""",1)"
            .Cells(6, 1).FormulaR1C1 = _
                "=IF(RC[1]<>1,"""",IF(RC[54]=VLOOKUP(RC[52],SF!C:C[11],12,FALSE),1,""""))"
            Range("A6:D" & EOLrep).FillDown
            
            PrevRep.Sheets(1).Rows(EOLprev - 2 & ":" & EOLprev).Copy .Cells(EOLrep + 1, 1)

        End With
    End With

    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.Path & "\WeeklySubsReport-03-Nov-2013.xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub
Function GetPrevRep() As Workbook
'
' - GetPrevRep() - открывает в каталоге Dir файл с предыдущим отчетом по SN
'
' 9.11.13

    Dim PrevRepName As String
    
    PrevRepName = Dir(ActiveWorkbook.Path & "\" & "WeeklySubsReport-*")
    If PrevRepName = ActiveWorkbook.Name Then
        PrevRepName = Dir()
    End If
    MsgBox "Прежний отчет = '" & PrevRepName & "'" _
        & vbCrLf & "  Новый отчет = '" & ActiveWorkbook.Name & "'"
    
    Set GetPrevRep = Workbooks.Open(ActiveWorkbook.Path & "\" & PrevRepName, , True)
End Function


