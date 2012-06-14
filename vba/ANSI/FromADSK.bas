Attribute VB_Name = "FromADSK"
'----------------------------------------------------------------------------
' ������ FromADSK - ������ � PartnerCenter.Autodesk.com
'
' * FrPartnerCenter()   - �������� ������ Subscription Renewal
' * wr3PASS()           - ������ ����� SN ��� 3PASS
' * rd3PASS()           - �������� outlook.csv ����� 3PASS
' * ACEreport()         - ������ ������ ACE
' T testGetSheetFrADSK()    - ���� �� ������ �������� TOC
' - GetSheetFrADSK(RepName) - ��������� ����� RepName �� ADSK.xlsx � Match
'
'   3.6.2012

    Option Explicit     ' Force explicit variable declaration
    
    Public Const PartnerCenter = "PartnerCenter"    ' ��� ����� ������ ��
                                '                  PartnerCenter.Autodesk.com
    Public Const A3PASS = "3PASS"       ' ���� ��� �������� 3PASS
    Public Const ACE = "ACE"            ' ���� ACE - ����������� �� SFDC/ADSK
    
 ' ����� ������� ������� ADSK
    Public Const StampADSKp1 = "Contract End Date"
    Public Const StampADSKp2 = "Days Until Contract Exp"
    Public Const ACEstamp = "Team Report"

Sub FrPartnerCenter()
'
' ??? �������� ������ �� CSV ����������� �� PartnerCenter.Autodesk.com
 ??? ��� ������ ����� ASFnm!!!!!!!!!!!!!!!!!!!!
'   27/1/2012

    Const ASFstamp = "Match ADSK from SF"   ' ����� - ��� ������ SF � �����
    Const Reslines = 6                      ' ���������� ����� ������ ����� ������� (�����)
    
    Const SheetNm = "PartnerCenter" ' �������� ����� - ������ ADSK � Match
    Const Astamp = "Renewal Name"   ' ����� ADSK � ������ I1 �������� �����
    Const ForCol = 7                ' ����� ������� - ������ �� ����� SheetNm
    Const ForColLtr = "G"
    
    Dim NewSheet As String          ' ��� ����� - ������ ������
    Dim LinesSF As Integer          ' ����� ����� � NewSheet - � ����� ������ ADSK
    
    Dim LO, Ln As Integer           ' ���������� ����� � ������ � ����� �������
    Dim Same As String
    
    LinesSF = ModStart(ASFnm, "���������� ������ Autodesk �� PartnerCenter")

    CheckSheet ASFnm, LinesSF - 4, 1, ASFstamp  ' ��������� ������� ������ �� SF
    CheckSheet SheetNm, 1, 16, Astamp           '   .. � �������� ������ �� PartnerCenter
    
    NewSheet = ADSKread()                           ' ������ Output.csv
    CheckSheet NewSheet, 1, 16 - ForCol, Astamp     '   .. � ��������� ��� ������������

    Sheets(SheetNm).Columns("A:" & ForColLtr).Copy  ' �������� ������� � ��������
    Sheets(NewSheet).Columns("A:A").Select          '   .. � ��������� � ����� �����
    Selection.Insert Shift:=xlToRight               '       .. �� ��������
    
' ��������� ������� ������ �� ����� ������� �������
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
    Sheets(SheetNm).Tab.Color = ADSK  ' ���������� Tab ������ ������

    ModEnd SheetNm
End Sub
Function ADSKread()
'
' ������� ������ ���� output.csv ����������� �� PartnerCenter.Autodesk.com
'
' ��� ��������� CSV �����
'   1. ������ � PartnerCenter.Autodesk.com
'   2. �� ������� <Subscription Renewals> Show <Contract Details Report> - <Export>
'   3. ��������� ��� ������� � ������� Tab Delimeted Text File �� ������� ����
'   4. ��������� ��� ���������� ������� <FrADSK> �� ����� "PartnerCenter"
'
'   27/1/2012

    ActiveWorkbook.Worksheets.add
    With ActiveSheet.QueryTables.add(Connection:= _
        "TEXT;C:\Users\������������\Desktop\output.csv", Destination:=Range( _
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
    Range("E:E, AG:AG").NumberFormat = "dd/mm/yy;@"         ' ������� - ����
    Range("Q:Q, AE:AE, AI:AI, AN:AN").NumberFormat = "@"    ' ������� - �����
    ADSKread = ActiveSheet.Name
End Function

Sub ACEreport()
'
' �������� ���� ACE ����� ������� �� PartnerCenter
'   22.4.2012
   
    ModStart ACE, "���������� ����� ACE - ������ Salesforce/Autodesk", True

    LinesOld = EOL(ACE) - SFresLines    ' ���-�� ����� � ������ ������
    Lines = EOL(1) - SFresLines         ' ���-�� ����� � ����� ������
    
    CheckSheet 1, Lines + 2, 1, ACEstamp
    CheckSheet ACE, LinesOld + 2, 4, ACEstamp
    
    Sheets(ACE).Columns("A:C").Copy     ' �� �������� ������ �������� ������� A:E
    Sheets(1).Columns("A:A").Select     '    � ��������� �� ����� � ������ ������
    Selection.Insert Shift:=xlToRight
    
    Range("F:I,K:I,M:P").Select           ' ������ ���������� �������� �������
    Selection.EntireColumn.Hidden = True
                                          
' ��������� ������� ������ �� ����� ������� �������
    If LinesOld < Lines Then
        Range(Cells(LinesOld, 1), Cells(LO, 5)).Select
        Selection.AutoFill Destination:=Range(Cells(LO, 1), Cells(Ln, 5)), _
            Type:=xlFillDefault
    Else
        Range(Cells(Lines + 1, 1), Cells(LinesOld + SFresLines, 3)).CleaContents
    End If
    
    Sheets(ACE).Select                  ' ��������� ������ ����� - ����� ������
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
' - GetSheetFrADSK(RepName) - ��������� ����� RepName �� ADSK
'       ���� ���� ���������� �� ADSK.xlsx � ����� RepName ��� ���� - ������� ��
'
'   27.05.12
'   3.6.12  - ������������ ��� ������, EOL_ADSKADSK_RepMap - Public
    
    Const TOC_RANGE_NAME = "TOC_RANGE_NAME"
    Const TOC_RANGE_MAP_OFFSET = 8  '�������� - ������� ��������
    
    Dim TOC_Line As Range
    Dim i As Integer
     
'---- ������� �������� ���������� �� ADSK.xlsx
    Workbooks.Open ("C:\Users\������������\Desktop\������ � Match\SFconstrTMP\ADSK\ADSK.xlsx")
    Windows("ADSK.xlsx").Activate
    Sheets(TOC_ADSK).Select
    Sheets(TOC_ADSK).Copy Before:=Workbooks("Match SF-1C.xlsm").Sheets(We)

'---- �������� TOC_Range � ������� � ��� ������ � Range ������� ������
    For Each TOC_Line In Range("TOC_ADSK_Range").Rows
        If TOC_Line.Cells(1, TOC_REPNAME_COL) = RepName Then Exit For
    Next TOC_Line

'---- ��������� Mapping ������� ������ �� ADSK.xlsx
    For i = 1 To ADSK_HdrMapSize
        ADSK_RepMap(i) = TOC_Line.Cells(1, TOC_RANGE_MAP_OFFSET + i)
    Next i
    
'---- ��������� ��� ����� - ������ ADSK � ��� ����� � Match
    Dim Rep() As String
    Rep = split(TOC_Line.Cells(1, TOC_REPRANGE_COL), "'")
    Workbooks("ADSK.xlsx").Sheets(Rep(0)).Copy Before:=Workbooks("Match SF-1C.xlsm").Sheets(We)
    Windows("ADSK.xlsx").Close
    Application.DisplayAlerts = False
    Workbooks("Match SF-1C.xlsm").Sheets(TOC_ADSK).Delete
    Workbooks("Match SF-1C.xlsm").Sheets(RepName).Delete
    Application.DisplayAlerts = True
    Workbooks("Match SF-1C.xlsm").Sheets(Rep(0)).Name = RepName
    Sheets(RepName).Tab.Color = rgbADSK  ' ���������� Tab ������ ������

'---- ��������� ��������� ������� ����������� ������
    For i = 1 To ADSK_HdrMapSize
        ADSK_HDR_Map(i) = Sheets(RepName).Cells(1, i)
    Next i
        
    ADSKrep = RepName
    EOL_ADSK = EOL(RepName)
End Sub
