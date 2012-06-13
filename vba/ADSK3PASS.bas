Attribute VB_Name = "ADSK3PASS"
'----------------------------------------------------------------------------
' Модуль ADSK3PASS - работа с PartnerCenter.Autodesk.com по Третьему Проходу
'
' название Третий проход (3PASS) сложилось исторически
' суть этой процедуры в головной программе SN3PASS:
' 1)                - очищаем лист 3PASS
' 2) SNselection    - готовим список SN из листа ADSKfrSF
' 3)   **           - из списка на листе 3PASS переносим (Copy/Paste) SN
'                     порциями на портал ADSK в Tab <License Inquiry>, затем <Go>
'                     эту подпрограмму запускаем по кнопке [2] многократно
'                     до исчерпания списка SN в А2 - начало порции, в А3 - конец
' 4) Load3passSN    - считываем на лист 3PASS выходные файлы из Autodesk
' 5) DoDeDupSN      - дедупликация серийных номеров на листе 3PASS
' 6) WrDL3pass      - запись из таблицы 3PASS в файл CSV для Data Loader
'
'   8.2.2012

    Option Explicit     ' Force explicit variable declaration
    
'    Public Const A3PASS = "3PASS"       ' лист для процедур 3PASS

    Dim Stage               ' стадия процесса 3PASS
    Private N As Integer    ' номер нижней строки с SN на листе 3PASS
    
Sub SN3pass()
'
' [*] - запуск 3PASS - третьего прохода, т.е. запись порций серийный номеров
'       в окно License Inquiry портала PartnerCenter по списку из отчета SF.
'       Затем считанные данные из PartnerCenter обрабатываются на листе 3PASS
'
'   8.2.2012

    Dim FrN, ToN, LastN ' начало и конец порции SN, последний SN
    Dim i, N As Integer ' номера порции, строки
    Const HelpStage2 = " скопирована в буфер клипборд." & vbCrLf & _
        vbCrLf & "Теперь:" & _
        vbCrLf & "1. в окне Serial Number <License Inquiry>" & _
        vbCrLf & "   PartnerCenter.Autodesk.com нажми Ctrl/V, затем [Go]" & _
        vbCrLf & "2. в Menu -> Export -> Next получи файлы output.csv." & _
        vbCrLf & _
        vbCrLf & "После появления output.csv в каталоге <Загрузки> жми [OK]."
                        
    Const SNstep = 100  ' максимальная порция SN
      
    Lines = Start3PASS("Cоcтавление списка SN")
    
    Rows("5:7777").Delete       ' очищаем прежнее содержимое листа 3PASS
    On Error Resume Next
    Kill DownloadDir & "output*.csv"    '   .. и файлы output*.csv
    On Error GoTo 0             ' на случай, когда output*.csv отсутствует
    
    SNselectionForm.Show
    Sheets(A3PASS).Select
    FrN = Cells(2, 1)
    ToN = Cells(3, 1)
    If FrN >= ToN Then End      ' если ни одного SN не выбрали - выход
     
'** получаем данные из PartnerCenter в файлы output.csv каталога "Загрузки"
    Start3PASS ("Считывание SN из PartnerCenter порциями в файлы output*.csv")
    LastN = ToN
    N = Round((LastN - FrN) / SNstep) + 1

    For i = 1 To N
        If ToN - FrN > SNstep Then ToN = FrN + SNstep
        Cells(ToN, 1) = Left(Cells(ToN, 1), 12) ' убираем последний +
        Range("A" & FrN & ":A" & ToN).Select    ' выбираем порцию
        Selection.Copy
        Selection.Interior.Color = rgbYellow    ' окрашиваем порцию желтым
        MsgBox i & "/" & N & ": SN от " & FrN & " до " & ToN & HelpStage2
        FrN = ToN + 1
        ToN = LastN
    Next i
    
    Load3passSN
'    DoDeDupSN
'    WrDL3pass
'    End3PASS 3
End Sub

Sub Load3passSN()
'
' [3] - чтение SN полученных из Licence Inquiry в файлов output.csv
' Файлов может быть несколько. После обработки они стираются.
'
'   1.2.2012

    Const CSVstamp1 = "Serial Number"
    Const CSVstamp2 = "Product Key"
    Dim F, SN, Status As String
    Dim R, C, i, j As Integer
    
    Start3PASS ("Загрузка файлов output*.csv")
        
' запись списка файлов output*.csv в колонку B листа 3PASS
    F = Dir(DownloadDir & "output*.csv")
    i = 5
    Do While F <> ""
        Cells(i, 2) = DownloadDir & F
        Cells(i, 2).Select
        F = Dir()
        i = i + 1
    Loop

' чтение файлов по списку в В
    R = 5: i = 5: C = 7
    Do While Cells(i, 2) <> ""
        Cells(i, 2).Select
        Call SNread(Cells(i, 2), R, C)    ' чтение файла output.csv
    ' проверка, правильный ли файл прочитали?
        If Cells(R, C) <> CSVstamp1 Or Cells(R, C + 1) <> CSVstamp2 Then
            MsgBox "Неправильный входной файл (" & F & _
                ") в листе 3PASS, начиная с ячейки (" _
                & R & "," & C & ")", vbCritical, "ERROR!"
            Stop
        End If
    ' удаление шапки прочитанного файла
        Range("G" & R & ":V" & R).Delete Shift:=xlUp
        While Cells(R, C) <> "" ' продвигаем R до нижней строки
            R = R + 1
        Wend
        i = i + 1
    Loop

    R = R - 1
    i = 5
    Do
        SN = Left(Cells(i, 1), 12)        ' желтый SN
'        Cells(i, 1).Select
        For j = 2 To R
            If SN = "" Then Exit Do
            If SN = Cells(j, 7) Then
                i = i + 1
                Exit For
            End If
        Next j
    Loop
        
    If i - 1 <> Cells(3, 1) Then
        MsgBox "Не все серийные номера нашлись -- см.лист 3PASS", , "ERROR!"
        Stop
    End If
        
    Range("B5:B" & i).ClearContents
    Range("A4:A" & Cells(3, 1)).ClearContents
    Range("B4:F" & R).FillDown  ' копируем формулы до конца (R)

    End3PASS 3
End Sub
Sub DoDeDupSN()
'
' [*] Дедупликация серийных номеров на листе 3PASS
'   6.2.2012

End Sub
Sub WrDL3pass()
'
' [3PASS] - запись при помощи Data Loader'a листа 3PASS в Salesforce
'   3/2/2012

    Start3PASS "Запись листа 3PASS посредством DataLoader в SF"
    
    Columns("C:C").Copy
    Columns("O:O").PasteSpecial Paste:=xlPasteValues
'    Columns("O:O").Select
'    Selection.PasteSpecial Paste:=xlPasteValues
    
    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\ADSK\3PASS\"
    WriteCSV A3PASS, "3PASS.txt", 1, 7

    Shell "quota2.bat 3PASS.TXT C:\SFconstr\3PASS.csv"
'    Shell "Copy 3PASS.csv C:\SFconstr"

    End3PASS (4)
End Sub
Function Start3PASS(Msg)
'
' начало процедур 3PASS. Возвращает номер последней значащей строки ADSKfrSF
'   2.2.2012
    Const txt = "Процедура 3PASS: получение SN из PartnerCenter --  "
    Const ASFstamp = "Match ADSK from SF"   ' штамп - имя Отчета SF в пятке
    Const A3PASSstamp = "SNfromSF"          ' штамп листа 3PASS
    Const Reslines = 6  ' число строк в пятке Отчета ADSKfrSF
    

    ModStart A3PASS, txt & Msg
    Start3PASS = EOL(ADSKfrSF) - Reslines
    CheckSheet ADSKfrSF, Start3PASS + 2, 1, ASFstamp
    CheckSheet A3PASS, 1, 1, A3PASSstamp
End Function
Sub End3PASS(M)
'
' завершение подпрограмм Stage3PASS
'   30/1/2012

    Const Colr1 = &HFFCC00  ' цвет [1] - голубой
    Const Colr2 = &HFF9900  ' цвет [2] - синий
    Const Colr3 = &HFF6600  ' цвет [1] - темно-синий
    Dim Colr
    
    Sheets(A3PASS).Select
    Select Case M
        Case 1
            Colr = Colr1
        Case 2
            Colr = Colr2
        Case 3
            Colr = Colr3
    End Select
    Cells(1, 1).Interior.Color = Colr   ' Окрашиваем ячейку А1
    Sheets(A3PASS).Tab.Color = Colr     '   .. и Tab 3PASS
    ModEnd A3PASS
End Sub
Sub SNread(FileNM, R, C)
'
' функция чтения файла серийных номеров FileNM на лист 3PASS в позицию (R,C)
'   31/1/2012

 '   FileNM = "C:\Users\Пользователь\Downloads\output.csv"
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & FileNM, _
        Destination:=Cells(R, C))
'        .Name = "output"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = _
            Array(2, 2, 2, 2, 2, 2, 2, 2, 3, 2, 2, 2, 2, 2, 2, 2)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub
