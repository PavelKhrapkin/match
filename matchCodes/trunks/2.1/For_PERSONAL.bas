Attribute VB_Name = "For_PERSONAL"
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
'   26.7.2012 - match 2.0 - MoveToMatch с использованием TOCmatch
'   17.8.2012 - Обработка Процессов - Loader'ов в ProcessEngine
'    8.9.2012 - этот модуль помещен под названием ForPERSONAL.bas, чтобы не путать

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
    If W Is Nothing Then
'------- нет, match.xlsm не там. Он есть среди Workbooks?
        For Each W In Application.Workbooks
            If W.Name = F Then
                P = W.Path & "\" & F
                GoTo RunMatch
            End If
        Next W
'------ и среди Workbooks нет. Посмотрим в файле в С:\
        On Error Resume Next
        Set W = Workbooks.Open(DinC)
        P = W.Sheets(1).Cells(1, 2) & F
        W.Close
        Set W = Workbooks.Open(P, UpdateLinks:=False)
        On Error GoTo 0
        If W Is Nothing Then
Const Msg = "<!> MoveToMatch не удалось открыть файл match.xlsm'" _
    & vbCrLf & vbCrLf & "Попробуй открыть его вручную, а потом" _
    & vbCrLf & "еще раз запусти MoveToMatch (Ctrl/ф)"
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
