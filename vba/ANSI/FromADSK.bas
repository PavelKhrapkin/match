Attribute VB_Name = "FromADSK"
'----------------------------------------------------------------------------
' Модуль FromADSK - работа с PartnerCenter.Autodesk.com
'
' * FrPartnerCenter()   - загрузка отчета Subscription Renewal
' * wr3PASS()           - запись файла SN для 3PASS
' * rd3PASS()           - загрузка outlook.csv после 3PASS
' * ACEreport()         - замена отчета ACE
' T testGetSheetFrADSK()    - пока на уровне загрузки TOC
' - GetSheetFrADSK(RepName) - загружает отчет RepName из ADSK.xlsx в Match
'
'   3.6.2012

    Option Explicit     ' Force explicit variable declaration
    
    Public Const PartnerCenter = "PartnerCenter"    ' имя листа отчета из
                                '                  PartnerCenter.Autodesk.com
    Public Const A3PASS = "3PASS"       ' лист для процедур 3PASS
    Public Const ACE = "ACE"            ' лист ACE - возможности из SFDC/ADSK
    
 ' имена штампов отчетов ADSK
    Public Const StampADSKp1 = "Contract End Date"
    Public Const StampADSKp2 = "Days Until Contract Exp"
    Public Const ACEstamp = "Team Report"

Sub FrPartnerCenter()
'
' ??? загрузка данных из CSV полученного из PartnerCenter.Autodesk.com
 ??? нет такого листа ASFnm!!!!!!!!!!!!!!!!!!!!
'   27/1/2012

    Const ASFstamp = "Match ADSK from SF"   ' штамп - имя Отчета SF в пятке
    Const Reslines = 6                      ' количество строк сводки после записей (пятка)
    
    Const SheetNm = "PartnerCenter" ' название листа - отчета ADSK в Match
    Const Astamp = "Renewal Name"   ' штамп ADSK в ячейке I1 входного файла
    Const ForCol = 7                ' число колонок - формул на листе SheetNm
    Const ForColLtr = "G"
    
    Dim NewSheet As String          ' имя листа - нового отчета
    Dim LinesSF As Integer          ' число строк в NewSheet - в новом отчете ADSK
    
    Dim LO, Ln As Integer           ' количества строк в старом и новом отчетах
    Dim Same As String
    
    LinesSF = ModStart(ASFnm, "Обновление отчета Autodesk из PartnerCenter")

    CheckSheet ASFnm, LinesSF - 4, 1, ASFstamp  ' проверяем наличие отчета из SF
    CheckSheet SheetNm, 1, 16, Astamp           '   .. и прежнего отчета из PartnerCenter
    
    NewSheet = ADSKread()                           ' читаем Output.csv
    CheckSheet NewSheet, 1, 16 - ForCol, Astamp     '   .. и проверяем его правильность

    Sheets(SheetNm).Columns("A:" & ForColLtr).Copy  ' копируем колонки с кнопками
    Sheets(NewSheet).Columns("A:A").Select          '   .. и формулами в новый отчет
    Selection.Insert Shift:=xlToRight               '       .. из прежнего
    
' дополняем колонки формул до конца рабочей области
    LO = EOL(SheetNm)
    Ln = EOL(NewSheet)

    If Ln > LO Then Range(Cells(LO, 1), Cells(Ln, ForCol)).FillDown

    Range("L1:L" & Ln).Interior.Color = RGB(152, 251, 152) ' Contract Start Date
    Range("AC1:AC" & Ln).Interior.Color = rgbYellow         ' Account #
    Range("AK1:AK" & Ln).Interior.Color = RGB(135, 206, 250) ' Serial Number
    Range("AM1:AM" & Ln).Interior.Color = RGB(154, 205, 50) ' Contract #
    Range("AN1:AN" & Ln).Interior.Color = RGB(107, 142, 35) ' Contract End Date
    
    Sheets(SheetNm).Delete
    Sheets(NewSheet).Name = SheetNm
    Sheets(SheetNm).Tab.Color = ADSK  ' окрашиваем Tab нового отчета

    ModEnd SheetNm
End Sub
Function ADSKread()
'
' функция чтения фала output.csv полученного из PartnerCenter.Autodesk.com
'
' Для получения CSV файла
'   1. Входим в PartnerCenter.Autodesk.com
'   2. на вкладке <Subscription Renewals> Show <Contract Details Report> - <Export>
'   3. выгружаем все колонки в формате Tab Delimeted Text File на рабочий стол
'   4. запускаем это приложение кнопкой <FrADSK> на листе "PartnerCenter"
'
'   27/1/2012

    ActiveWorkbook.Worksheets.add
    With ActiveSheet.QueryTables.add(Connection:= _
        "TEXT;C:\Users\Пользователь\Desktop\output.csv", Destination:=Range( _
        "$A$1"))
'        .Name = "output (1)_1"
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
        .TextFileColumnDataTypes = Array(2, 2, 2, 2, 3, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 3, 2, 2, _
        2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 2, 3, 1, 1, 1, 1, 1, 1, 1, 1, 1, 3)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Range("E:E, AG:AG").NumberFormat = "dd/mm/yy;@"         ' колонки - даты
    Range("Q:Q, AE:AE, AI:AI, AN:AN").NumberFormat = "@"    ' колонки - числа
    ADSKread = ActiveSheet.Name
End Function

Sub ACEreport()
'
' Заменяет лист ACE новым отчетом из PartnerCenter
'   22.4.2012
   
    ModStart ACE, "Обновление листа ACE - отчета Salesforce/Autodesk", True

    LinesOld = EOL(ACE) - SFresLines    ' кол-во строк в старом отчете
    Lines = EOL(1) - SFresLines         ' кол-во строк в новом отчете
    
    CheckSheet 1, Lines + 2, 1, ACEstamp
    CheckSheet ACE, LinesOld + 2, 4, ACEstamp
    
    Sheets(ACE).Columns("A:C").Copy     ' из прежнего отчета копируем колонки A:E
    Sheets(1).Columns("A:A").Select     '    и вставляем их слева к новому отчету
    Selection.Insert Shift:=xlToRight
    
    Range("F:I,K:I,M:P").Select           ' делаем невидимыми ненужные колонки
    Selection.EntireColumn.Hidden = True
                                          
' дополняем колонки формул до конца рабочей области
    If LinesOld < Lines Then
        Range(Cells(LinesOld, 1), Cells(LO, 5)).Select
        Selection.AutoFill Destination:=Range(Cells(LO, 1), Cells(Ln, 5)), _
            Type:=xlFillDefault
    Else
        Range(Cells(Lines + 1, 1), Cells(LinesOld + SFresLines, 3)).CleaContents
    End If
    
    Sheets(ACE).Select                  ' вставляем группу ячеек - итоги сверки
    Range(Cells(LinesOld - 2, 2), Cells(LinesOld, 18)).Copy
    Sheets(1).Select
    Range(Cells(Lines - 1, 2), Cells(Lines - 1, 2)).Activate
    ActiveSheet.Paste

    For i = 2 To Lines
        AceN = Cells(i, ACE_N_COL)
        if instr(
    Next i
End Sub
Sub testGetSheetFrADSK()
    GetSheetFrADSK "GFP"
End Sub
Sub GetSheetFrADSK(RepName)
'
' - GetSheetFrADSK(RepName) - загружает отчет RepName из ADSK
'       если лист Оглавление из ADSK.xlsx и отчет RepName уже есть - стираем их
'
'   27.05.12
'   3.6.12  - возвращаемые имя отчета, EOL_ADSKADSK_RepMap - Public
    
    Const TOC_RANGE_NAME = "TOC_RANGE_NAME"
    Const TOC_RANGE_MAP_OFFSET = 8  'смещение - колонка маппинга
    
    Dim TOC_Line As Range
    Dim i As Integer
     
'---- вначале загружем Оглавление из ADSK.xlsx
    Workbooks.Open ("C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\ADSK\ADSK.xlsx")
    Windows("ADSK.xlsx").Activate
    Sheets(TOC_ADSK).Select
    Sheets(TOC_ADSK).Copy Before:=Workbooks("Match SF-1C.xlsm").Sheets(We)

'---- получаем TOC_Range и находим в нем строку с Range нужного отчета
    For Each TOC_Line In Range("TOC_ADSK_Range").Rows
        If TOC_Line.Cells(1, TOC_REPNAME_COL) = RepName Then Exit For
    Next TOC_Line

'---- извлекаем Mapping колонок отчета из ADSK.xlsx
    For i = 1 To ADSK_HdrMapSize
        ADSK_RepMap(i) = TOC_Line.Cells(1, TOC_RANGE_MAP_OFFSET + i)
    Next i
    
'---- извлекаем имя листа - отчета ADSK и сам отчет в Match
    Dim Rep() As String
    Rep = split(TOC_Line.Cells(1, TOC_REPRANGE_COL), "'")
    Workbooks("ADSK.xlsx").Sheets(Rep(0)).Copy Before:=Workbooks("Match SF-1C.xlsm").Sheets(We)
    Windows("ADSK.xlsx").Close
    Application.DisplayAlerts = False
    Workbooks("Match SF-1C.xlsm").Sheets(TOC_ADSK).Delete
    Workbooks("Match SF-1C.xlsm").Sheets(RepName).Delete
    Application.DisplayAlerts = True
    Workbooks("Match SF-1C.xlsm").Sheets(Rep(0)).Name = RepName
    Sheets(RepName).Tab.Color = rgbADSK  ' окрашиваем Tab нового отчета

'---- извлекаем заголовки колонок полученного отчета
    For i = 1 To ADSK_HdrMapSize
        ADSK_HDR_Map(i) = Sheets(RepName).Cells(1, i)
    Next i
        
    ADSKrep = RepName
    EOL_ADSK = EOL(RepName)
End Sub
