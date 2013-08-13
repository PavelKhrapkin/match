Attribute VB_Name = "ADSKPartnerCenterPass"
'----------------------------------------------------------------------------
' Модуль ADSKPartnerCenterPass - работа с License Inquiry PartnerCenter.Autodesk.com
'
''' суть этой процедуры в головной программе SN3PASS:
''' 1)                - очищаем лист 3PASS
''' 2) SNselection    - готовим список SN из листа ADSKfrSF
''' 3)   **           - из списка на листе 3PASS переносим (Copy/Paste) SN
'''                     порциями на портал ADSK в Tab <License Inquiry>, затем <Go>
'''                     эту подпрограмму запускаем по кнопке [2] многократно
'''                     до исчерпания списка SN в А2 - начало порции, в А3 - конец
''' 4) Load3passSN    - считываем на лист 3PASS выходные файлы из Autodesk
''' 5) DoDeDupSN      - дедупликация серийных номеров на листе 3PASS
''' 6) WrDL3pass      - запись из таблицы 3PASS в файл CSV для Data Loader
'
'   8.2.2012
'  13.8.2013 - переписано для match 2.0

    Option Explicit     ' Force explicit variable declaration
    
    Dim Stage               ' стадия процесса 3PASS
    Private N As Integer    ' номер нижней строки с SN на листе 3PASS
    
Sub SN_PC_pass()
'
' S SN_PC_Pass - получение серийных номеров (SN) из PartnerCenter.Autodesk.com
'                по списку в SN_LIST по порциям, а затем сортировка этих данных
'                на SN_ACTIVE и SN_UPDATE.
'
'   13.8.2013

    Const SN_LIST = "A_PC_1"
    Const SN_TMP = "A_PC_2"
    Const ST_ACTIVE = "A_PC_3"
    Const ST_UPDATE = "A_PC_4"
    
    Dim R As TOCmatch
    Dim i, N As Integer ' номера порции, строки
    Dim FrN, ToN, LastN ' начало и конец порции SN, последний SN
    Const HelpStage = " скопированы в буфер клипборд." & vbCrLf & _
        vbCrLf & "Теперь:" & _
        vbCrLf & "1. в окне Serial Number <License Inquiry>" & _
        vbCrLf & "   PartnerCenter.Autodesk.com нажми Ctrl/V, затем [Go]" & _
        vbCrLf & "2. в Menu -> Export -> Next получи файлы output.csv." & _
        vbCrLf & _
        vbCrLf & "После появления output.csv в каталоге <Загрузки> жми [OK]."
                        
    Const SNstep = 100  ' максимальная порция SN
    
    StepIn
    R = GetRep(SN_LIST)
    MS "Проверяем в PartnerCenter.Autodesk.com " & R.EOL & " строк Registered SN"
    
 '---------- инициализация -------------------
    iActive = 1: iUpdate = 1
    nActive = 0: nUpdate = 0: nToClean = 0
    
    Dim Rtmp As TOCmatch, Ractive As TOCmatch, Rupdate As TOCmatch
    
    Rtmp = GetRep(SNtmp)
    Ractive = GetRep(SNactive)
    Rupdate = GetRep(SNupdate)
    NewSheet Ractive.Name
    NewSheet Rupdate.Name
    
  '--------------- цикл по SN_List ---------------
    i = 2
    With Workbooks(R.RepFile).Sheets(R.SheetN)
        .Activate
        Do While i <= R.EOL
            ToN = i + SNstep
            If ToN > R.EOL Then ToN = R.EOL
            .Cells(ToN, 1) = Left(.Cells(ToN, 1), 12)   ' убираем последний +
            Range(.Cells(i, 1), .Cells(ToN, 1)).Select  ' выбираем порцию
            Selection.Copy
            Selection.Interior.Color = rgbYellow        ' окрашиваем порцию желтым
            
            On Error Resume Next
            Kill DownloadDir & "output*.csv"    ' Очищаем все старые файлы output*.csv
            On Error GoTo 0                     ' на случай, когда output*.csv отсутствует
            
            MsgBox "SN от " & i & " до " & ToN & HelpStage
            SNread SN_TMP
            SNsortOut SN_TMP, SN_ACTIVE, SN_UPDATE
        Loop
    End With
 End Sub
Sub SNsortOut(SNtmp, SNactive, SNupdate)
'
' - SNsortOut(SNtmp, SNactive, SNupdate) - разбор и селекция SN из SNtmp.
'      - в SNactive заносятся данные о Registered SN
'      - в SNudpade - об Upgraded SN, которые надо изменить в SF
'   Кроме того в SNtmp изменяется цвет SN:
'           - белый      - SN найден, он Registered
'           - коричневый - найден, требуется Update
'           - остается желтый - не найден или требует ручной проверки
'
'   13.8.2013

    Dim iTmp As Long    '- указатель - номер строки в SNtmp
    Call SheetSort(Rtmp.Name, 1)
    
    For iTmp = 2 To Rtmp.EOL
    
    Next iTmp
End Sub

Sub LoadSNfrOutputCSV()
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
    F = dir(DownloadDir & "output*.csv")
    i = 2
    Do While F <> ""
        Cells(i, 2) = DownloadDir & F
        Cells(i, 2).Select
        F = dir()
        i = i + 1
    Loop

' чтение файлов по списку в В
    R = 5: i = 2: C = 7
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

    ' проверяем, все ли SN из списка в колонке А нашлись?
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

    Shell "quota_3PASS.bat"
'    Shell "Copy 3PASS.csv C:\SFconstr"

    End3PASS (4)
End Sub
Function Start3PASS(Msg)
'
' начало процедур 3PASS. Возвращает номер последней значащей строки ADSKfrSF
'   2.2.2012
    Const txt = "Процедура 3PASS: получение SN из PartnerCenter --  "
    Const A3PASSstamp = "SNfromSF"          ' штамп листа 3PASS
    Const ResLines = 6  ' число строк в пятке Отчета ADSKfrSF
    

    ModStart A3PASS, txt & Msg
    Start3PASS = EOL(ADSKfrSF) - ResLines
    CheckSheet ADSKfrSF, Start3PASS + 2, 1, ADSKfrSFstamp
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
Sub SNread(F)
'
' - SNread  - функция чтения файла серийных номеров из Output.CSV
'             в верхний левый угол листа F
'   31/1/2012
'   13.08.13 - упрощен интерфейс - чтение всегда в A_SN_2

    Dim R As TOCmatch

    NewSheet F      ' Сбрасываем файл SN из PartnerCenter
    
    R = GetRep(F)
    If R.EOL <> 1 Then ErrMsg FATAL_ERR, "Лист \\W_TMP.A_PC_2' не очищен!"
    Workbooks(R.RepFile).Sheets(R.SheetN).Activate
    
    With ActiveSheet.QueryTables.Add( _
            Connection:="TEXT;" & DownloadDir & "output.csv", _
            Destination:=Cells(1, 1))
        .Name = "output"
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
        On Error GoTo Rep
        .Refresh BackgroundQuery:=False
    End With
    Exit Sub
Rep:
    If MsgBox("В каталоге 'Загрузки' не найден файл 'output.csv'." _
        & vbCrLf & "'" & DownloadDir & "output.csv" & "'" _
        & vbCrLf & vbCrLf & "Убедись, что он загружен из PartnerCenter и повтори.") _
        = vbYes Then
            Call SNread
    Else
            End
    End If
End Sub
