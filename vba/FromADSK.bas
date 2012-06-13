п»їAttribute VB_Name = "FromADSK"
'----------------------------------------------------------------------------
' РњРѕРґСѓР»СЊ FromADSK - СЂР°Р±РѕС‚Р° СЃ PartnerCenter.Autodesk.com
'
' * FrPartnerCenter()   - Р·Р°РіСЂСѓР·РєР° РѕС‚С‡РµС‚Р° Subscription Renewal
' * wr3PASS()           - Р·Р°РїРёСЃСЊ С„Р°Р№Р»Р° SN РґР»СЏ 3PASS
' * rd3PASS()           - Р·Р°РіСЂСѓР·РєР° outlook.csv РїРѕСЃР»Рµ 3PASS
' * ACEreport()         - Р·Р°РјРµРЅР° РѕС‚С‡РµС‚Р° ACE
' T testGetSheetFrADSK()    - РїРѕРєР° РЅР° СѓСЂРѕРІРЅРµ Р·Р°РіСЂСѓР·РєРё TOC
' - GetSheetFrADSK(RepName) - Р·Р°РіСЂСѓР¶Р°РµС‚ РѕС‚С‡РµС‚ RepName РёР· ADSK.xlsx РІ Match
'
'   3.6.2012

    Option Explicit     ' Force explicit variable declaration
    
    Public Const PartnerCenter = "PartnerCenter"    ' РёРјСЏ Р»РёСЃС‚Р° РѕС‚С‡РµС‚Р° РёР·
                                '                  PartnerCenter.Autodesk.com
    Public Const A3PASS = "3PASS"       ' Р»РёСЃС‚ РґР»СЏ РїСЂРѕС†РµРґСѓСЂ 3PASS
    Public Const ACE = "ACE"            ' Р»РёСЃС‚ ACE - РІРѕР·РјРѕР¶РЅРѕСЃС‚Рё РёР· SFDC/ADSK
    
 ' РёРјРµРЅР° С€С‚Р°РјРїРѕРІ РѕС‚С‡РµС‚РѕРІ ADSK
    Public Const StampADSKp1 = "Contract End Date"
    Public Const StampADSKp2 = "Days Until Contract Exp"
    Public Const ACEstamp = "Team Report"

Sub FrPartnerCenter()
'
' Р·Р°РіСЂСѓР·РєР° РґР°РЅРЅС‹С… РёР· CSV РїРѕР»СѓС‡РµРЅРЅРѕРіРѕ РёР· PartnerCenter.Autodesk.com
'   27/1/2012

    Const ASFstamp = "Match ADSK from SF"   ' С€С‚Р°РјРї - РёРјСЏ РћС‚С‡РµС‚Р° SF РІ РїСЏС‚РєРµ
    Const Reslines = 6                      ' РєРѕР»РёС‡РµСЃС‚РІРѕ СЃС‚СЂРѕРє СЃРІРѕРґРєРё РїРѕСЃР»Рµ Р·Р°РїРёСЃРµР№ (РїСЏС‚РєР°)
    
    Const SheetNm = "PartnerCenter" ' РЅР°Р·РІР°РЅРёРµ Р»РёСЃС‚Р° - РѕС‚С‡РµС‚Р° ADSK РІ Match
    Const Astamp = "Renewal Name"   ' С€С‚Р°РјРї ADSK РІ СЏС‡РµР№РєРµ I1 РІС…РѕРґРЅРѕРіРѕ С„Р°Р№Р»Р°
    Const ForCol = 7                ' С‡РёСЃР»Рѕ РєРѕР»РѕРЅРѕРє - С„РѕСЂРјСѓР» РЅР° Р»РёСЃС‚Рµ SheetNm
    Const ForColLtr = "G"
    
    Dim NewSheet As String          ' РёРјСЏ Р»РёСЃС‚Р° - РЅРѕРІРѕРіРѕ РѕС‚С‡РµС‚Р°
    Dim LinesSF As Integer          ' С‡РёСЃР»Рѕ СЃС‚СЂРѕРє РІ NewSheet - РІ РЅРѕРІРѕРј РѕС‚С‡РµС‚Рµ ADSK
    
    Dim LO, Ln As Integer           ' РєРѕР»РёС‡РµСЃС‚РІР° СЃС‚СЂРѕРє РІ СЃС‚Р°СЂРѕРј Рё РЅРѕРІРѕРј РѕС‚С‡РµС‚Р°С…
    Dim Same As String
    
    LinesSF = ModStart(ASFnm, "РћР±РЅРѕРІР»РµРЅРёРµ РѕС‚С‡РµС‚Р° Autodesk РёР· PartnerCenter")

    CheckSheet ASFnm, LinesSF - 4, 1, ASFstamp  ' РїСЂРѕРІРµСЂСЏРµРј РЅР°Р»РёС‡РёРµ РѕС‚С‡РµС‚Р° РёР· SF
    CheckSheet SheetNm, 1, 16, Astamp           '   .. Рё РїСЂРµР¶РЅРµРіРѕ РѕС‚С‡РµС‚Р° РёР· PartnerCenter
    
    NewSheet = ADSKread()                           ' С‡РёС‚Р°РµРј Output.csv
    CheckSheet NewSheet, 1, 16 - ForCol, Astamp     '   .. Рё РїСЂРѕРІРµСЂСЏРµРј РµРіРѕ РїСЂР°РІРёР»СЊРЅРѕСЃС‚СЊ

    Sheets(SheetNm).Columns("A:" & ForColLtr).Copy  ' РєРѕРїРёСЂСѓРµРј РєРѕР»РѕРЅРєРё СЃ РєРЅРѕРїРєР°РјРё
    Sheets(NewSheet).Columns("A:A").Select          '   .. Рё С„РѕСЂРјСѓР»Р°РјРё РІ РЅРѕРІС‹Р№ РѕС‚С‡РµС‚
    Selection.Insert Shift:=xlToRight               '       .. РёР· РїСЂРµР¶РЅРµРіРѕ
    
' РґРѕРїРѕР»РЅСЏРµРј РєРѕР»РѕРЅРєРё С„РѕСЂРјСѓР» РґРѕ РєРѕРЅС†Р° СЂР°Р±РѕС‡РµР№ РѕР±Р»Р°СЃС‚Рё
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
    Sheets(SheetNm).Tab.Color = ADSK  ' РѕРєСЂР°С€РёРІР°РµРј Tab РЅРѕРІРѕРіРѕ РѕС‚С‡РµС‚Р°

    ModEnd SheetNm
End Sub
Function ADSKread()
'
' С„СѓРЅРєС†РёСЏ С‡С‚РµРЅРёСЏ С„Р°Р»Р° output.csv РїРѕР»СѓС‡РµРЅРЅРѕРіРѕ РёР· PartnerCenter.Autodesk.com
'
' Р”Р»СЏ РїРѕР»СѓС‡РµРЅРёСЏ CSV С„Р°Р№Р»Р°
'   1. Р’С…РѕРґРёРј РІ PartnerCenter.Autodesk.com
'   2. РЅР° РІРєР»Р°РґРєРµ <Subscription Renewals> Show <Contract Details Report> - <Export>
'   3. РІС‹РіСЂСѓР¶Р°РµРј РІСЃРµ РєРѕР»РѕРЅРєРё РІ С„РѕСЂРјР°С‚Рµ Tab Delimeted Text File РЅР° СЂР°Р±РѕС‡РёР№ СЃС‚РѕР»
'   4. Р·Р°РїСѓСЃРєР°РµРј СЌС‚Рѕ РїСЂРёР»РѕР¶РµРЅРёРµ РєРЅРѕРїРєРѕР№ <FrADSK> РЅР° Р»РёСЃС‚Рµ "PartnerCenter"
'
'   27/1/2012

    ActiveWorkbook.Worksheets.add
    With ActiveSheet.QueryTables.add(Connection:= _
        "TEXT;C:\Users\РџРѕР»СЊР·РѕРІР°С‚РµР»СЊ\Desktop\output.csv", Destination:=Range( _
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
    Range("E:E, AG:AG").NumberFormat = "dd/mm/yy;@"         ' РєРѕР»РѕРЅРєРё - РґР°С‚С‹
    Range("Q:Q, AE:AE, AI:AI, AN:AN").NumberFormat = "@"    ' РєРѕР»РѕРЅРєРё - С‡РёСЃР»Р°
    ADSKread = ActiveSheet.Name
End Function

Sub ACEreport()
'
' Р—Р°РјРµРЅСЏРµС‚ Р»РёСЃС‚ ACE РЅРѕРІС‹Рј РѕС‚С‡РµС‚РѕРј РёР· PartnerCenter
'   22.4.2012
   
    ModStart ACE, "РћР±РЅРѕРІР»РµРЅРёРµ Р»РёСЃС‚Р° ACE - РѕС‚С‡РµС‚Р° Salesforce/Autodesk", True

    LinesOld = EOL(ACE) - SFresLines    ' РєРѕР»-РІРѕ СЃС‚СЂРѕРє РІ СЃС‚Р°СЂРѕРј РѕС‚С‡РµС‚Рµ
    Lines = EOL(1) - SFresLines         ' РєРѕР»-РІРѕ СЃС‚СЂРѕРє РІ РЅРѕРІРѕРј РѕС‚С‡РµС‚Рµ
    
    CheckSheet 1, Lines + 2, 1, ACEstamp
    CheckSheet ACE, LinesOld + 2, 4, ACEstamp
    
    Sheets(ACE).Columns("A:C").Copy     ' РёР· РїСЂРµР¶РЅРµРіРѕ РѕС‚С‡РµС‚Р° РєРѕРїРёСЂСѓРµРј РєРѕР»РѕРЅРєРё A:E
    Sheets(1).Columns("A:A").Select     '    Рё РІСЃС‚Р°РІР»СЏРµРј РёС… СЃР»РµРІР° Рє РЅРѕРІРѕРјСѓ РѕС‚С‡РµС‚Сѓ
    Selection.Insert Shift:=xlToRight
    
    Range("F:I,K:I,M:P").Select           ' РґРµР»Р°РµРј РЅРµРІРёРґРёРјС‹РјРё РЅРµРЅСѓР¶РЅС‹Рµ РєРѕР»РѕРЅРєРё
    Selection.EntireColumn.Hidden = True
                                          
' РґРѕРїРѕР»РЅСЏРµРј РєРѕР»РѕРЅРєРё С„РѕСЂРјСѓР» РґРѕ РєРѕРЅС†Р° СЂР°Р±РѕС‡РµР№ РѕР±Р»Р°СЃС‚Рё
    If LinesOld < Lines Then
        Range(Cells(LinesOld, 1), Cells(LO, 5)).Select
        Selection.AutoFill Destination:=Range(Cells(LO, 1), Cells(Ln, 5)), _
            Type:=xlFillDefault
    Else
        Range(Cells(Lines + 1, 1), Cells(LinesOld + SFresLines, 3)).CleaContents
    End If
    
    Sheets(ACE).Select                  ' РІСЃС‚Р°РІР»СЏРµРј РіСЂСѓРїРїСѓ СЏС‡РµРµРє - РёС‚РѕРіРё СЃРІРµСЂРєРё
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
' - GetSheetFrADSK(RepName) - Р·Р°РіСЂСѓР¶Р°РµС‚ РѕС‚С‡РµС‚ RepName РёР· ADSK
'
'   27.05.12
'   3.6.12  - РІРѕР·РІСЂР°С‰Р°РµРјС‹Рµ РёРјСЏ РѕС‚С‡РµС‚Р°, EOL_ADSKADSK_RepMap - Public

'----- РћРіР»Р°РІР»РµРЅРёРµ Р±Р°Р·С‹ ADSK.xlsx --------------------
    Const TOC_ADSK = "TOC_ADSK"
    Const TOC_REPNAME_COL = 3   'РїРѕР»Рµ - РРјСЏ/С‚РёРї РѕС‚С‡РµС‚Р°
    Const TOC_REPRANGE_COL = 5  'РїРѕР»Рµ - Р›РёСЃС‚ (Range)
    
    Const TOC_RANGE_NAME = "TOC_RANGE_NAME"
    Const TOC_RANGE_MAP_OFFSET = 9  'СЃРјРµС‰РµРЅРёРµ - РєРѕР»РѕРЅРєР° РјР°РїРїРёРЅРіР°
    
    Dim TOC_Line As Range
    Dim i As Integer
    
'---- РµСЃР»Рё Р»РёСЃС‚ РћРіР»Р°РІР»РµРЅРёРµ РёР· ADSK.xlsx Рё РѕС‚С‡РµС‚ RepName СѓР¶Рµ РµСЃС‚СЊ - СЃС‚РёСЂР°РµРј РёС…
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(TOC_ADSK).Delete
    Sheets(RepName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
'---- РІРЅР°С‡Р°Р»Рµ Р·Р°РіСЂСѓР¶РµРј РћРіР»Р°РІР»РµРЅРёРµ РёР· ADSK.xlsx
    Workbooks.Open ("C:\Users\РџРѕР»СЊР·РѕРІР°С‚РµР»СЊ\Desktop\Р Р°Р±РѕС‚Р° СЃ Match\SFconstrTMP\ADSK\ADSK.xlsx")
    Windows("ADSK.xlsx").Activate
    Sheets(TOC_ADSK).Select
    Sheets(TOC_ADSK).Copy Before:=Workbooks("Match SF-1C.xlsm").Sheets(We)

'---- РїРѕР»СѓС‡Р°РµРј TOC_Range Рё РЅР°С…РѕРґРёРј РІ РЅРµРј СЃС‚СЂРѕРєСѓ СЃ Range РЅСѓР¶РЅРѕРіРѕ РѕС‚С‡РµС‚Р°
    For Each TOC_Line In Range("TOC_ADSK_Range").Rows
        If TOC_Line.Cells(1, TOC_REPNAME_COL) = RepName Then Exit For
    Next TOC_Line

'---- РёР·РІР»РµРєР°РµРј Mapping РєРѕР»РѕРЅРѕРє РѕС‚С‡РµС‚Р° РёР· ADSK.xlsx
    For i = 1 To ADSK_HdrMapSize
        ADSK_RepMap(i) = TOC_Line.Cells(1, TOC_RANGE_MAP_OFFSET + i)
    Next i
    
'---- РёР·РІР»РµРєР°РµРј РёРјСЏ Р»РёСЃС‚Р° - РѕС‚С‡РµС‚Р° ADSK Рё СЃР°Рј РѕС‚С‡РµС‚ РІ Match
    Dim Rep() As String
    Rep = split(TOC_Line.Cells(1, TOC_REPRANGE_COL), "'")
    Workbooks("ADSK.xlsx").Sheets(Rep(0)).Copy Before:=Workbooks("Match SF-1C.xlsm").Sheets(We)
    Windows("ADSK.xlsx").Close
    Application.DisplayAlerts = False
    Workbooks("Match SF-1C.xlsm").Sheets(TOC_ADSK).Delete
    Application.DisplayAlerts = True
    Workbooks("Match SF-1C.xlsm").Sheets(Rep(0)).Name = RepName
    Sheets(RepName).Tab.Color = rgbADSK  ' РѕРєСЂР°С€РёРІР°РµРј Tab РЅРѕРІРѕРіРѕ РѕС‚С‡РµС‚Р°

'---- РёР·РІР»РµРєР°РµРј Р·Р°РіРѕР»РѕРІРєРё РєРѕР»РѕРЅРѕРє РїРѕР»СѓС‡РµРЅРЅРѕРіРѕ РѕС‚С‡РµС‚Р°
    For i = 1 To ADSK_HdrMapSize
        ADSK_HDR_Map(i) = Sheets(RepName).Cells(1, i)
    Next i
        
    ADSKrep = RepName
    EOL_ADSK = EOL(RepName)
End Sub
