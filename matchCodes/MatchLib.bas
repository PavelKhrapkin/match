Attribute VB_Name = "MatchLib"
'---------------------------------------------------------------------------
' ���������� ����������� ������� "match 2.1"
'
' �.�.�������, �.���� 8.9.13
'
' S setColWidth(file, sheet, col, range, width) - ������������� ������ ������� �����
' S InsMyCol()                  - ��������� ������� MyCol � ���� ����� � ����� �� ��������
' - MS(Msg)                     - ����� ��������� �� ����� � � LogWr
' - ErrMsg(ErrMode, MSG)        - ����� ��������� �� ������ � Log � �� �����
' - LogWr(msg)                  - ������ ��������� msg � Log list
'(*)LogReset()                  - ����� � ������� Log �����
' - AutoFilterReset(SheetN)     - ����� � ����������� ����������� ����� SheetN
' - SheetsCtrlH(SheetN, FromStr, ToStr) - ������ ������ FromStr �� ToStr
'                                 � ����� SheetN
' - PerCent(Row, Col)           - �������������� ������ (Row,Col) � ����������
' - CurCode(Row, Col, CurCol)   - ������ ������ (Row,Col) �� ���� ������ � CurCol
' - CurRate(Cur)                - ���������� ���� ������ � ����� �� ���� Cur ��� We
' - CurISO(Cur1C)               - ���������� ��� ������ � ��������� ISO
' - DDMMYYYY(d)                 - �������������� ���� d � ��������� ������ DDMMYYYY
' - GetDate(txt)                - ��������� ���� �� ��������� ������ txt
' - Dec(a)                      - ������ ����� � � ���� ������ � ���������� ������
' - EOL(SheetN)                 - ���������� ����� ��������� ������ ����� SheetN
' - RowDel(RowStr)              - ������� ������ ��������� ����� � ������������ � RowStr
' T testCSmatch()               - ������� CSmatch
' - CSmatch(Val,Col)            - Case Sensitive match - ���������� ����� ������ � Val
'                                 � ������� Col. ���� Val �� ������- ���������� 0.
'                                 ���� ��� ������ Val ������ ���� Selected.
' - CSmatchSht(Val,Col,Sht)     - Case Sensitive match ���������� ����� ������ � Val �
'                                 ������� Col ����� Sht. ���� Val �� ������- ���������� 0.
' - SheetExists(SheetName)      - ��������, ��� ���� SheetName ��������
'[X]ClearSheet(SheetN, HDR_Range) - ������� ����� SheetN � ������ � ���� �����
' - SheetSort(SheetN, Col)      - ���������� ����� SheetN �� ������� Col
' - SheetDedup(SheetN, Col)     - c��������� � ������������ SheetN �� ������� Col
' - SheetDedup2(SheetN, ColSort,ColAcc) - ���������� � ������� ����� SheetN
'                                 �� �������� ColSort, ColAcc
' - DateCol(SheetN, Col)        - �������������� ������� Col �� ������ � ����
' - DateSort(Col)               - �������������� ������� Col ����� ActiveSheet
'                                 �� ���������� ������� � Date
'                                 � ���������� �� ���� ������� �� ������ � ����� �����
' - HideLns(FrR, ToR, Col, Criteria) - �������� ������ �� FrR �� ToR,
'                                 ���� Col �����.Criteria (��� WP)
' - Progress(Pct)               - ����� Progress Bar'� - ��������� ���������� ����
' - StopSub()                   - ��������� ���������� �������� - ���������� �� �����
' - RemDelimeters(S)            - ������ ���� �������� � ������ S �� �������
' - Compressor(S)               - �������� ������� �������� �� ������ S
' � testFindInLst()             - ������� FindInLst(W,Lst)
' - IsInList(W,Lst)             - ���������� ���� �� ����� �� ������ W � ������ Lst
' � testFindInDIC()             - ������� FindInDIC(W,Dic)
' - IsInDIC(W,Dic)              - ���������� ���� �� ����� W � ������� Dic
' - IsMatchList(W, DicList)     - ���������� TRUE ���� W ������ � DicList
' T test ISML                   - ������� IsMatchList
' � testpatTest()               - ������� patTest
' - patTest                     - �������� �� ������������ ����������� ���������

Option Explicit

' ��������� - ������������ ������
    Public Const rgbADSK = 2162853  ' ����������� ������-���������� ���� ��� Autodesk
    Public Const LimeG = 3329330    ' RGB(50, 205, 50)  �������-�������
    Public Const Antique = 11583693 ' RGB(205, 192, 176) ����� - ����-����������
    Public Const Gray = 8750469     ' �����
    
    Public Const Log = "Log"        ' Log ����
    
' ���������� � �������� �������� '������'
    Dim patObject
    Dim patObjectSet As Boolean
        
Sub InsMyCol()
'
' S InsMyCol() - ��������� ������� MyCol � ���� ����� � ����� �� �������� � ���
'
'   * ���� ��������� ������� ������� ����� ������ - ����������
'   * ���� � ������ 2 ����� ������� "V" - ������������ ����� �� �������
'
'  10.8.12
'  15.8.12 - Optional FS
'  26.8.12 - RowHeight ����� ��� � �������; ���� ������ 2 "V" - �������� �����
'  31.8.12 - ��������� StepIn
'  11.9.12 - ������� ���� � Headers ����� match.xlsm
'  1.10.12 - ����������� ��������� ������� � ����� �� COPY_HDR � ������ 2 �������
'  4.11.12 - ������������� R=GetRep(ActiveSheet.Name)
' 19.11.12 - COPY_HDR - copy ������� ������ ������������ ���������� �����������
' 19.12.12 - ��������� ������� � ������ Width
' 27.12.12 - ����������� ��� ������ ������
' 12.01.13 - ������������� ������ �����o� �������� Application.DecimalSeparator
' 13.01.13 - ������� ������� ������� �� setColWidth
' 20.01.13 - ���� �� ������� ������� ������� �� 1 ��������
' 28.01.13 - width � setColWidth ������ ������: ������/������
' 24.08.13 - ����� �������� �������� �� TOC

    Const COPY_HDR = "CopyHdr"

    StepIn
    
    Dim R As TOCmatch   'R - ��������� TOCmatch ��� SFD
    Dim FF As Range, Fsummary As String
    Dim i As Integer
'''    Set FF = DB_MATCH.Sheets(Header).Range(F)
    
    R = GetRep(ActiveSheet.Name)
    
    With DB_MATCH.Sheets(TOC)
        Set FF = DB_MATCH.Sheets(Header).Range(.Cells(R.iTOC, TOC_FORMNAME))
        Fsummary = .Cells(R.iTOC, TOC_FORMSUMMARY)
    End With
    
    With Workbooks(R.RepFile).Sheets(R.SheetN)
        .Activate
'---- � ����� �� ��� ��� ������� ���������?
        If .Cells(1, 1) = FF.Cells(1, 1) Then Exit Sub

'---- ��������� ������� �� ����� MyCol
        .Range(Cells(1, 1), Cells(1, R.MyCol)).EntireColumn.Insert
'---- ������ ������ � ��������� ����������� �������
        For i = 1 To FF.Columns.Count
            setColWidth R.RepFile, R.SheetN, i, FF.Cells(3, i)
            If FF.Cells(2, i) = COPY_HDR Then
                FF.Cells(1, i).Copy Destination:=.Cells(1, i)
            End If
        Next i
'---- �������� ������� MyCol �� ����� �� EOL
        For i = 1 To R.MyCol
            FF.Cells(1, i).Copy Destination:=.Cells(1, i)
            FF.Cells(2, i).Copy Destination:=.Cells(2, i)
        Next i
        .Rows(1).RowHeight = FF.Rows(1).RowHeight
        .Range(.Cells(2, 1), .Cells(R.EOL, R.MyCol)).FillDown
'---- ��������� ����� �� ������� � Fsummary
        If Fsummary = "" Then Exit Sub
        Set FF = DB_MATCH.Sheets(Header).Range(Fsummary)
        FF.Copy Destination:=.Cells(R.EOL + 2, 1)
    End With
End Sub
Sub testsetColWidth()
' � testsetColWidth() - ������� setColWidth
'12.1.13

    Set DB_MATCH = FileOpen(F_MATCH)
    Dim FF As Range
    
    Set FF = DB_MATCH.Sheets(Header).Range("HDR_1C_Payment_MyCol")
    Dim i As Long
a:
    For i = 1 To FF.Columns.Count
        setColWidth "1C.xlsx", "�������", i, FF.Cells(3, i)
    Next i
        
    Stop
    GoTo a

End Sub
Sub setColWidth(ByVal file As String, ByVal SHEET As String, _
                ByVal Col As Long, ByVal width As String)
'
' - setColWidth(file, sheet, col, range, width) -
'           ������������� ������ i-� ������� �����
' 12.01.2013
' 16.1.13 bug fix
' 16.1.13 ��������� �� ����� ������ ���������� ������� �� ����� (��� ��������)
' 28.1.13 width ������ ������: ������/������
'
    Dim widSplit() As String
    widSplit = Split(width, "/")
    If UBound(widSplit) >= 0 Then
        width = widSplit(0)
        If Application.DecimalSeparator = "." Then
            width = Replace(width, ",", ".")        ' ������������ ������ - �������� ',' �� '.'
        ElseIf Application.DecimalSeparator = "," Then
            width = Replace(width, ".", ",")        ' ���������� ������ - �������� '.' �� ','
        End If
        Workbooks(file).Sheets(SHEET).Columns(Col).ColumnWidth = CSng(width)
    End If

End Sub
Sub MS(msg)
'
'   - MS(Msg)- ����� ��������� �� ����� � � LogWr
'   11.6.12
    ErrMsg TYPE_ERR, msg
End Sub

Sub ErrMsg(ErrMode, msg)
'
' - ErrMsg(ErrMode, MSG) - ����� ��������� �� ������ � Log � �� �����
'                          ���� ErrMode ���������� � Declaration
'   31.5.12

    Dim ErrType As String, Respond As String

    Select Case ErrMode
    Case WARNING:
        LogWr "< WARNING > " & msg
        Exit Sub
        
    Case TYPE_ERR:
        LogWr "��������:" & msg
        Respond = MsgBox(msg & vbCrLf & vbCrLf & "����������?", vbYesNo)
        If Respond = vbNo Then
            ExRespond = False
            Stop
        End If
        Exit Sub
        
    Case FATAL_ERR:
Fatal:  ErrType = "<! ERROR !> "
        LogWr ErrType & msg
        MsgBox msg, , ErrType
        Stop
        Exit Sub
    Case Else:
        ErrMode = FATAL_ERR
        GoTo Fatal
    End Select
End Sub
Sub LogWr(msg)
'
' ������ ��������� msg � Log-����
'   15.2.2012
'   26.6.12 - match 2.0
'    9.9.12 - ������ ��������� Log � ������ ������ ����� ����� � Log match.xlsm
'   17.7.13 - Workbooks(F_MATCH) ������ DB_MATCH

    Dim N   ' ����� ������ � Log
    
    With Workbooks(F_MATCH).Sheets(Log)
        N = .Cells(1, 4)
        N = N + 1
        .Cells(N, 1) = Date
        .Cells(N, 2) = Time
        .Cells(N, 3) = msg
        .Cells(1, 4) = N
    End With
End Sub
Sub LogReset()
'
' ����� Log ����� � ��� �������
' ����������� ������� [Reset] �� ����� Log
'   19.2.2012

    Dim N
    Sheets(Log).Select
    Cells(1, 3).Select
    N = Cells(1, 4)
    Cells(1, 4) = 1
    Rows("2:" & N).Delete
    LogWr ("LogReset")
End Sub
Function AutoFilterReset(SheetN) As Integer
'
' - AutoFilterReset(SheetN)   - ������������ ������ � ������ �������
'                               ����� SheetN � ������ ������
'                               ���������� ���������� ����� � SheetN
' 16.1.2012
' 15.12.12 - ���������� ��� �������� �� MS Office 2013:
'       - EOL � ������������ � ���������� match2.0 ������� �� TOCmatch
'       - AutoFilter ��������������� �� EOL, ����� ���������� �����
'  25.8.13 - ����������� FreezePanes � Scroll �� �����

    Dim R As TOCmatch
    
    R = GetRep(SheetN)
    
    With Workbooks(R.RepFile).Sheets(R.SheetN)
        .Activate
        .AutoFilterMode = False
        .Rows("1:" & R.EOL).AutoFilter
        .Rows("1:1").Select
        On Error Resume Next
        ActiveWindow.FreezePanes = True
        Application.GoTo .Cells(R.EOL - 3, 1), True
        On Error GoTo 0
    End With
End Function
Sub tst()
'    AutoFilterReset PAY_SHEET
    Const SHEET_N = 1
    AutoFilterReset SHEET_N
End Sub

Sub SheetsCtrlH(SheetN, FromStr, ToStr)
'
' ������������ ������ (Ctrl/H) ������ FromStr �� ToStr � ����� SheetN
'   7.1.2012
'  27.1.2012 - ����� Activate/Select
    
    Call AutoFilterReset(SheetN)

    Cells.Replace What:=FromStr, Replacement:=ToStr, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
Sub PerCent(Row, Col)
'
' ������������ �������������� %
'   26.1.12

    Cells(Row, Col).NumberFormat = "@"
    Cells(Row, Col) = Cells(Row, Col) & "%"
End Sub
Sub CurCode(Row, Col, CurCol)
'
' ������������ �������������� ����� � ������ (Row,Col).
' ��� ������ � ��� �� ������ � ������� Col
'   20.1.12

    Select Case Cells(Row, CurCol)
        Case "RUB"
            Cells(Row, Col).NumberFormat = "_-* #,##0.00[$�.-419]_-;-* #,##0.00[$�.-419]_-;_-* ""-""??[$�.-419]_-;_-@_-"
        Case "EUR"
            Cells(Row, Col).NumberFormat = "_-[$�-2] * #,##0.00_ ;_-[$�-2] * -#,##0.00 ;_-[$�-2] * ""-""??_ ;_-@_ "
        Case "USD"
            Cells(Row, Col).NumberFormat = "_-[$$-409]* #,##0.00_ ;_-[$$-409]* -#,##0.00 ;_-[$$-409]* ""-""??_ ;_-@_ "
        Case Else
            MsgBox "ERROR in WPopp: �������� ��� ������ = " & Cells(Row, CurCol), , "ERROR!"
    End Select
End Sub
Function CurRate(Cur) As Double
'
' ���������� ����� - ���� � ����� �� ���� ������ Cur �� ������� Currence �� ����� We
'   21.2.2012
'   20.8.12 - ������������� "���"
'    4.9.12 - ��������� Sheets(We)

    Dim S

    CurRate = 1
    If InStr(LCase(Cur), "���") > 0 Or Trim(Cur) = "" Then Exit Function

    On Error GoTo Col2
    S = WorksheetFunction.VLookup(Cur, DB_MATCH.Sheets(We).Range("RUB_Rate"), 3, False)
    On Error GoTo 0
    GoTo Convert
Col2:
    On Error GoTo 0
    S = WorksheetFunction.VLookup(Cur, DB_MATCH.Sheets(We).Range("RUB_Rate_2"), 2, False)
Convert:
    CurRate = Replace(S, ".", ",")
End Function
Function CurISO(Cur1C)
'
' ���������� ��� ������ � ��������� ISO, ������������ ��� �� ���� 1�
'   18.3.2012
'    4.9.2012 - ��������� Sheets(We)
'   19.9.12 - �� ��������� CurISO="RUB"

    CurISO = "RUB"
    On Error Resume Next
    CurISO = WorksheetFunction.VLookup(Cur1C, DB_MATCH.Sheets(We).Range("Currency"), 2, False)
    On Error GoTo 0
End Function
Function DDMMYYYY(D) As String
'
' �������������� ���� d � ��������� ������ DDMMYYYY
'   14.2.2012
    DDMMYYYY = Day(D) & "." & Month(D) & "." & Year(D)
End Function
Function GetDate(txt As String) As Date
'
' - GetDate(txt) - �������������� ������ txt � ����
' 24.12.12

    Dim componentArray() As String, new_txt As String
    If IsDate(txt) Then
        GetDate = txt           ' �������� ������ ����������������
    Else
            ' ������������ '��.��.���� �����' -> '��/��/���� �����'
            
        componentArray = Split(txt, ".")
        new_txt = componentArray(1) & "/" & componentArray(0) & "/" & componentArray(2)
        If Not IsDate(new_txt) Then
            ' �� ����. ������������ '��.��.���� �����' -> '��/��/���� �����'
            
            componentArray = Split(txt, "/")
            new_txt = componentArray(1) & "." & componentArray(0) & "." & componentArray(2)
            If Not IsDate(new_txt) Then ErrMsg FATAL_ERR, "GetDate - ������������ ������ ����"
        End If
        GetDate = new_txt
    End If

End Function
Sub tGetDate()

    Dim res(1 To 5) As Date
    res(1) = GetDate("12/24/2012 4:12")
    res(2) = GetDate("12.24.2012 4:12")
    res(3) = GetDate("24.2.12 4:12")
    Stop
End Sub
Function Dec(a) As String
'
' �������������� ����� � � ��������� ������ � ���������� ������
'   14.2.2012

    Dec = "'" & WorksheetFunction.Substitute(a, ",", ".")
End Function
Sub testEOL()
    Dim a, b, C
    a = EOL(1)
    b = EOL(2)
        Dim F As Workbook
        Set F = Workbooks.Open(F_SFDC, UpdateLinks:=True, ReadOnly:=True)
    C = EOL(1, F)
        F.Close SaveChanges:=False
End Sub
Function EOL(ByVal SheetN As String, Optional F As Workbook = Nothing)
'
' - EOL(SheetN,[F]) - ���������� ���������� ����� � ����� SheetN ����� F
'   20/1/2012
'   4/2/2012 - ��������� ������ On Error
'   20/2/2012 - ��������� Option Explicit
'   12.5.12 - Sheets(SheetN).Select ��������
'   24.6.12 - AllCol - Public
'   29.6.12 - match 2.0 - ������� ���� S, ���� ������ ��������
'   31.7.12 - ���� ActiveWorkbook � �� ThisWorkbook �� ���������,
'             ByVal SheetN As String - �������������� �������� ���������
'   20.8.12 - �� �������������� SheetN ������������ EOL = -1

    Dim i

    If F Is Nothing Then
        Set F = ActiveWorkbook
    End If
    
    EOL = -1
    On Error Resume Next
    EOL = F.Sheets(SheetN).UsedRange.Rows.Count
    On Error GoTo 0
    If EOL <= 0 Then Exit Function
    
    With F.Sheets(SheetN)
        AllCol = .UsedRange.Columns.Count
        Do
            For i = 1 To AllCol
                If .Cells(EOL, i) <> "" Then Exit Do
            Next i
            If EOL <= 1 Then Exit Do
            EOL = EOL - 1       ' ������ UsedRange ��������� ������ ������,
        Loop                    '   .. ��������, ���� � ������ ���� ��������� ������
    End With
End Function
Sub RowDel(RowStr As String)
'
' - RowDel(RowStr) - ������� ������ ��������� ����� � ������������ � RowStr
'   25.8.12
    StepIn
    ActiveSheet.Rows(RowStr).Delete
End Sub
Sub testCSmatch()
    If "G" = "g" Then Stop
    Dim a
    ThisWorkbook.Sheets("Sheet1").Select
    a = CSmatch("g12", 1)
    a = CSmatch("g121", 1)
    
    ModStart REP_1C_P_PAINT
    Set DB_SFDC = Workbooks.Open(F_SFDC, UpdateLinks:=False, ReadOnly:=True)
    DB_SFDC.Sheets(SFacc).Select
    a = CSmatch("��� ""���""", 2)
    ModEnd
End Sub
Function CSmatch(Val, Col) As Double
'
' - CSmatch(Val,Col) - Case Sensitive match ���������� ����� ������ � Val � ������� Col.
'                   ���� Val �� ������- ���������� 0. ���� ��� ������ Val ������ ���� Selected.
' 8/7/12

    Dim CheckCS
    Dim N As Long
    N = 1
    Do
        CSmatch = 0
        On Error Resume Next
        CSmatch = Application.Match(Val, Range(Cells(N, Col), Cells(BIG, Col)), 0) + N - 1
        CheckCS = Cells(CSmatch, Col)
        On Error GoTo 0
        If IsEmpty(CSmatch) Or Not IsNumeric(CSmatch) Or CSmatch <= 0 Then Exit Function
        N = CSmatch + 1
    Loop While Val <> CheckCS
End Function
Function CSmatchSht(Val, Col, sht, Optional ByVal FromN As Long = 1) As Long
'
' - CSmatch(Val,Col,Sht) - Case Sensitive match ���������� ����� ������ � Val � ������� Col.
'                   ���� Val �� ������- ���������� 0. Sht - ���� ��� ������ Val.
' 27.9.12
' 25.11.12 - Optional FromN
'  5.12.12 - bug fix

    Dim CheckCS
''    Dim N As Long
''    N = 1
    Do
        CSmatchSht = 0
        On Error Resume Next
        CSmatchSht = Application.Match(Val, Range(sht.Cells(FromN, Col), sht.Cells(BIG, Col)), 0) _
            + FromN - 1
        CheckCS = sht.Cells(CSmatchSht, Col)
        On Error GoTo 0
''        If IsEmpty(CSmatchSht) Or Not IsNumeric(CSmatchSht) Or CSmatchSht <= 0 Then Exit Function
        If Not IsNumeric(CSmatchSht) Or CSmatchSht <= 0 Then Exit Function
        FromN = CSmatchSht + 1
    Loop While Val <> CheckCS
End Function
Function SheetExists(SheetName As String) As Boolean
'
' - SheetExists(SheetName)  - ��������, ��� ���� SheetName ��������
'
' 18.8.13 �� ���������

    On Error Resume Next
    SheetExists = Not Sheets(SheetName) Is Nothing
End Function
Sub ClearSheet(SheetN, HDR_Range As Range)
'
' ������ ������� SheetN � ������� � ���� ��������� �� ����� �eader.HDR_Range
'   4.2.2012
'  11.2.2012 - ��������� ������������
'  10.3.12 - ��������� ������������ - �������� HRD_Range
'  25.3.12 - ����� NewContract � NewContractLnk
'  17.4.12 - ���� A_Acc - ����� �����������
'  18.4.12 - ���� A_Dic - ������� �����������
'  28.4.12 - ���� NewOrderList - ���� ����� �������
'  13.5.12 - ���� P_ADSKlink - ����� ������ ������ - ADSK
'  15.5.12 - ���� SF_PA ������ �������� � ����������� ADSK
'   6.6.12 - Delete ������ ����, ������� �����
'  11.6.12 - ����� A_Acc � AccntUpd
'  12.6.12 - ���� BTO_SHEET - ��� ��� ����� ���

    DB_MATCH.Sheets(SheetN).Activate
    
' -- ������� ������ ����
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(SheetN).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
' -- ������� ����� ����
    Sheets.Add After:=Sheets(Sheets.Count)  ' ������� ����� ���� � ����� ������
    ActiveSheet.Name = SheetN
    ActiveSheet.Tab.Color = RGB(50, 153, 204)   ' Tab �������
   
    HDR_Range.Copy Sheets(SheetN).Cells(1, 1)   ' �������� ����� �� Header
    
    Select Case SheetN
    Case O_NewOpp:      EOL_NewOpp = 1
    Case P_Paid:        EOL_NewPay = 1
    Case NewContract:       EOL_NewContr = 1
    Case NewContractLnk:    EOL_ContrLnk = 1
    Case P_PaymentUpd:  EOL_PaymentUpd = 1
    Case A_Dic:         EOL_DIC = 1
    Case A_Acc:         EOL_AdAcc = 1
    Case AccntUpd:      EOL_AccntUpd = 1
    Case NewOrderList:  EOL_NewOrderList = 1
    Case P_ADSKlink:    EOL_ADSKlnkPay = 1
    Case SF_PA:         EOL_SFlnkADSK = 1
    Case NewSN:         EOL_NewSN = 1
    Case BTO_SHEET:     EOL_BTO = 1
'    Case WP:        EOL_WP = 1
    Case Else
        MsgBox "ClearSheet: ������� ��������� ����� '" & SheetN & "'" _
            , , "ERROR!"
        Stop
    End Select
End Sub
Sub SheetSort(SheetN, Col, Optional AscOrder As String = "Ascending")
'
' -/S SheetSort(SheetN, Col,[AscOrder]) - ��������� ���� SheetN �� ������� Col
'                                    � ������� AscOrder (�� ��������� Ascending)
'       * ���������, ������� �� AutoFilter � ����������� �����.
'         ���� �������, �� ��������� �� Range AutoFilter, ����� - �� ����� �����.
'
'   22.1.2012
'   21.2.2012 - Option Explicit
'   19.4.12 - AutoFilterReset
'   21.8.13 - ���� ��������� ����� ������� ��������- �� �������� AutoFilterReset
'   1.9.13 - ������� ���������� - Optional AscOrder
'   8.9.13 - ���������� ������ ���������� ��� AutoFilterResert

    Dim Name As String, Ord
    
    If Not IsNumeric(SheetN) Then Call AutoFilterReset(SheetN)

    Name = ActiveSheet.Name
    
    Ord = xlAscending
    If UCase$(Left$(AscOrder, 3)) = "DES" Then Ord = xlDescending
        
    With ActiveWorkbook.Worksheets(Name).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Cells(1, Col), SortOn:=xlSortOnValues, _
            Order:=Ord, DataOption:=xlSortNormal
        If ActiveSheet.AutoFilter Is Nothing Then
            .SetRange Range("A1:ZZ99999")    '!! �� ������ ������
        Else
            .SetRange ActiveSheet.AutoFilter.Range
        End If
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub SheetDedup(SheetN, Col)
'
' ������� ������ - ��������� � ���� SheetN �� ������� Col,
'                  �������� ���������� �� ���� �������
'   19.4.2012

    Dim i, prev, x, EOL_SheetN As Integer
    
    Call SheetSort(SheetN, Col)
    EOL_SheetN = EOL(SheetN)
    
    prev = "": i = 2
    Do
        x = Sheets(SheetN).Cells(i, Col)
        If x = prev Then
            Rows(i).Delete
            EOL_SheetN = EOL_SheetN - 1
        Else
            i = i + 1
            prev = x
        End If
    Loop While i < EOL_SheetN
End Sub
Sub SheetDedup2(SheetN, ColSort, �olAcc, ColIdSF)
'
' - SheetDedup2(SheetN, ColSort, ColAcc, ColIdSF)  - ��������� ���� SheetN
'          �� ������� SortCol ����� �������� ColAcc � ColIdSF ������ � "+"
'   23.5.2012
'   23.11.12 - ������� � match2.0

    Dim i As Integer, EOL_SheetN As Integer
    Dim prev As String, x As String
    Dim PrevAcc As String, NewAcc As String
    Dim PrevSFid As String, NewSFid As String
    
    Call SheetSort(SheetN, ColSort)
    EOL_SheetN = EOL(SheetN)
    
    prev = "": i = 2
    With Sheets(SheetN)
        Do
            x = .Cells(i, ColSort)
            If x = prev Then
                PrevAcc = .Cells(i - 1, �olAcc)
                PrevSFid = .Cells(i - 1, ColIdSF)
                NewAcc = .Cells(i, �olAcc)
                NewSFid = .Cells(i, ColIdSF)
                If PrevAcc <> "" And NewAcc <> "" And PrevAcc <> NewAcc Then
                    PrevAcc = PrevAcc & "+" & NewAcc
                ElseIf PrevAcc = "" And NewAcc <> "" Then
                    PrevAcc = NewAcc
'                ElseIf PrevAcc <> "" And NewAcc = "" Then
'                ElseIf PrevAcc = "" And NewAcc = "" Then
'                   � ���� ��������� ������� ������ �� ������
                End If
                If PrevSFid <> "" And NewSFid <> "" And PrevSFid <> NewSFid Then
                    .Cells(i - 1, ColIdSF) = PrevSFid & "+" & NewSFid
                End If
                .Cells(i - 1, �olAcc) = PrevAcc
                Rows(i).Delete
                EOL_SheetN = EOL_SheetN - 1
            Else
                i = i + 1
                prev = x
            End If
        Loop While i < EOL_SheetN
    End With
End Sub
Sub DateCol(ByVal SheetN As String, ByVal Col As Integer)
'
' �������������� ������� Col � ����� SheetN �� ������ ���� DD.MM.YY � ������ Date
'   20.4.12
'   3.10.12 - GetRep ������ EOL
'   2.1.13  - ������� 2000 � ����, ���� �� ����������, � ������������� � ���������� �������
'   3.1.13  - �������� ������ ����
'  21.8.13  - ��� MoveInMatch

    Dim i As Long, dd As Long, MM As Long, YY As Long
    Dim Dat As Date
    Dim D() As String
    
    Dim R As TOCmatch
    If Not IsNumeric(SheetN) Then
        R = GetRep(SheetN)
        Workbooks(R.RepFile).Sheets(R.SheetN).Activate
        Lines = R.EOL
    End If
    
    With ActiveSheet
        For i = 2 To Lines
            D = Split(.Cells(i, Col), ".")
            If UBound(D) = 2 And IsNumeric(D(0)) And IsNumeric(D(1)) And IsNumeric(D(2)) Then
                dd = D(0)
                If dd < 1 Or dd > 31 Then GoTo Nxt
                MM = D(1)
                If MM < 1 Or MM > 12 Then GoTo Nxt
                YY = D(2)
                If YY < 100 Then YY = 2000 + YY
                Dat = GetDate(dd & "." & MM & "." & YY)
                .Cells(i, Col) = Dat
            End If
Nxt:    Next i
    End With
End Sub
Sub DateSort(ByVal Col As Integer)
'
' S DateSort(Col) - �������������� ������� Col �� ���������� ������� � Date
'                   � ���������� �� ���� ������� �� ������ � ����� �����
'   31.7.12
'   31.8.12 - �������� ��� Step �� StepIn
'   25.8.13 - ��� MoveToMatch ����� DateCol � Sheetsort ������ ������.
'             SheetN �������� �� ActiveSheet.Name

    StepIn
    DateCol ActiveSheet.Name, Col
    SheetSort ActiveSheet.Name, Col
End Sub
Sub HideLns(FrR, ToR, Col, _
    Optional Criteria As String, Optional HideFlag As Boolean = True)
'
' �������� (Hide) ������ �����, ��������������� Criteria
' ��������������� ���� � ������� Col �� ������ FrR �� ToR
' - ���� �������� HideFlag=FALSE, �� �������� ������ Ungide
' - ���� Col = 0 - Hide ��� Unhide ��� ������ �� FrR �� ToR
'   25.1.2012
'    4.2.2012 - Unhide ����� �� HideFlag=FALSE

    Dim RowsList As String      ' ������ ����� ��� Hide
    Dim RowsListLng As Integer  ' ����� ������ �����
    
    RowsList = ""
    RowsListLng = 0
    If Col = 0 Then
       RowsList = FrR & ":" & ToR
       RowListLng = 1
    Else
        For i = FrR To ToR
            If Cells(i, Col) = Criteria Then
                If RowsListLng > 0 Then RowsList = RowsList & ","
                RowsList = RowsList & i & ":" & i
                RowsListLng = RowsListLng + 1
            End If
        Next i
        If RowsListLng = 0 Then Exit Sub    ' ���� ������ ������ �� ����
    End If
    Range(RowsList).Select      ' �������:  Range("30:30,31:31")
    Selection.EntireRow.Hidden = HideFlag
End Sub
Sub Progress(Pct)
'
' ����� Progress Bar'� - ��������� ���������� ����
'   15/2/2012
'   26.5.12 - MsgBox ������ 20 ��� �� ����� ���������� Progress
'   29.5.12 - ��������� �����
'    7.8.12 - ������� �������� ����� �� StatusBar
'   31.8.12 - ��������� StepIn
            
    Application.StatusBar = PublicProcName & "> " _
        & "��� " & PublicStepName _
        & ": " & Format(Pct, "0%")
    End Sub
Sub StopSub()
'
' StopSub() ��������� ���������� �������� - ���������� �� ������� FATAL ERROR
'
    
    MsgBox "��������� ��������� StopSub", , "FATAL ERROR"
    Stop
End Sub
Function RemDelimiters(S)
'
' RemDelimeters(S) - ������ ���� �������� � ������ S �� �������
' 19.4.12 �.����

    Dim smb, i

    For i = 1 To Len(S)
        
        ' ��������� �����, ��������� ��� ������� �����
        smb = Mid(S, i, 1)
        If Not ((smb >= "0" And smb <= "9") _
                Or (smb >= "A" And smb <= "Z") _
                Or (smb >= "a" And smb <= "z") _
                Or (smb >= "�" And smb <= "�") _
                Or (smb >= "�" And smb <= "�") _
                Or smb = "�" Or smb = "�" _
                Or smb = "�" Or smb = "�") Then
            Mid(S, i, 1) = " "      ' ������ ����������, �������� �� ������
        End If
    Next i
    RemDelimiters = S
End Function
Function Compressor(S1 As Variant)
'
' �������� ������ �������� ������ ������
'   7.3.12  �� ��������
'   7.6.12 ������� vbCcLf

    Dim S As Variant
    S = Replace(S1, vbCrLf, " ")
    S = Trim(S)
    While InStr(1, S, "  ") <> 0
       S = Left(S, InStr(1, S, "  ") - 1) & Right(S, Len(S) - InStr(1, S, "  "))
    Wend
    Compressor = S
End Function
Sub testFindInLst()
'
' � testFindInLst() - ������� FindInLst(W,Lst)
'   24/5/12
    Dim a
    a = FindInLst("Autodesk Plant Design Suite Premium 2012 New SLM", "DIC_GoodADSK")
End Sub
Function FindInLst(W, Lst) As String
'
' - IsInList(W,Lst)  - ���������� ���� �� ����� �� ������ W � ������ Lst
'   24.5.12

    Dim S As Range
    Dim lW As String, V As String
    
    lW = LCase$(W)
    For Each S In Range(Lst).Rows
        V = S.Cells(1, 1)
        If InStr(lW, LCase$(V)) <> 0 Then
            FindInLst = V
            Exit For
        End If
    Next S
End Function
Sub testFindInDIC()
'
' � testFindInDIC() - ������� FindInDIC(W,Dic)
'   7/5/12
    Dim a
    a = FindInDIC("��������", "Goods")
End Sub
Function FindInDIC(W, Dic) As String
'
' - IsInDIC(W,Dic)  - ���������� ���� �� ����� W � ������� Dic
'       ������������� ���������� �� ������ ������� Dic
'       � ��� ����� ��� �������������, ����������� ��������.
'       ���� �� ������ ������� ����� - ���������� �� ������ � ������
'   7.5.12

    FindInDIC = ""
    Dim S As Range
    Dim ValWord As String
    Dim Article As String

    For Each S In Range(Dic).Rows
        ValWord = S.Cells(1, 1)
        Article = S.Cells(1, 2)
        If Article = "" Then Article = ValWord
        If IsMatchList(W, Article) Then
            FindInDIC = ValWord
            Exit For
        End If
    Next S
End Function
Function IsMatchList(W, DicList) As Boolean
'
' - IsMatchList(W, DicList) - ���������� TRUE ���� W ������ � DicList
'   7.5.12

    IsMatchList = False
    If W = "" Or DicList = "" Then Exit Function
    
    Dim x() As String
    Dim i As Integer
    Dim lW As String
    
    lW = LCase$(W)
    x = Split(DicList, ",")
    
    For i = LBound(x) To UBound(x)
        If InStr(lW, LCase$(x(i))) <> 0 Then
            IsMatchList = True
            Exit Function
        End If
    Next i
End Function
Sub testISML()
'
' T test ISML - ������� IsMatchList
' 7/5/12
    Dim a As Boolean
    a = IsMatchList("", "�����,�����,���")
    a = IsMatchList("������", "�����,�����,���,���")
    a = IsMatchList("������", "�����,�����,���")
End Sub
Sub ScreenUpdate(TurnOn As Boolean)
'
' - ScreenUpdate(ToDo) - switch off Screen Update if TurnOn = False
'   8.11.12

    If Not TurnOn Then
        With Application
            .ScreenUpdating = False
'            .EnableEvents = False
            .DisplayAlerts = False
        End With
    Else
        With Application
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
            .DisplayStatusBar = True
            .DisplayAlerts = True
        End With
    End If
End Sub
Sub testpatTest()

    Dim ret(1 To 15) As Boolean
'                               ���                 ���
a:
    ret(1) = patTest("xxx-TEST-yyyccc", "TEST")                    '��
    ret(2) = patTest("xxx-TEST-yyyccc", "^TEST$")      '��� - �.�. �������
    ret(3) = patTest("xxx-TEST-yyyccc", "^xxx-TEST-yyyccc$")       '��
    ret(4) = patTest("xxx-TzST-yyyccc", "^xxx-T.ST-yyyccc")        '��
    ret(5) = patTest("xxx-TEST-yyyccc", "T[eE]ST")                 '��
    ret(6) = patTest("xxx-TeST-yyyccc", "^xxx-T[eE]ST-yyyccc$")    '��
    ret(7) = patTest("xxx-TeST-yyyccc", "^xxx-T\wST-yyyccc$")      '��
    ret(8) = patTest(" xxx �������5 ; ��� xxx ������� ", "$���.*�������;�������(\s|\d|$)")  '��
    ret(9) = patTest(" xxx �������a ; ��� xxx ������� ", "$���.*�������;�������(\s|\d|$)")  '���(2)
    ret(10) = patTest(" xxx �������x; ��� xxx ������� ", "$��x�.*�������;�������(\s|$)")    '��
    ret(11) = patTest(" xxx �������x; ��� xxx ������� ", "$���.*��������;�������(\s|$)")    '��
    ret(12) = patTest(" xxx ������� ; ��� xxx ������� ", "$���.*�������;�������(\s|\d|$)")  '��
    ret(12) = patTest("����������� xxx �������� ; ��� xxx ��� ", "$�����������;�������")  '���
    Stop
    GoTo a
    
End Sub

Function patTest(longTxt As String, pat As String) As Boolean
'
' - patTest - �������� �� ������������ ����������� ���������
'             ���������� True, ���� ������ longTxt ������������� ������� pat
'   22.12.2012
'   28.12.12 - replace "~" with ","
'   29.12.12 - ������� ����������� $pat1;pat0 ��� pat0 - ������� �������;
'                   $pat1; - ����� �����������, ����� ��� longTxt ������ ��� �������������.
'               ������: $���*.printer;printer( |\d|$). ����� '��� printer' �������� ����,
'                       � '��� ��� printer' - �� ��������.
'
    
    patTest = False
    If Not patObjectSet Then
        Set patObject = CreateObject("VBSCRIPT.REGEXP")
        patObjectSet = True
    End If
    
    With patObject
         
        If Left(pat, 1) = "$" Then
            pat = Trim(Mid(pat, 2))             ' ������� ����. ������
            If pat = "" Then GoTo Ex
            
            Dim comps() As String, pats() As String, i As Long
                ' � ������ 2-�� ����� '$' - ���������� �������� ������, ����������� ��������,
                '       ��������������� ����������
                '       � � �������� ������ ���� ����� 2 �����
            pats = Split(pat, ";")
            If UBound(pats) <> 1 Then GoTo Ex  ' � �������� ������ ���� ����� 2 �����
            
            comps = Split(longTxt, ";")
            For i = 0 To UBound(comps)
                    ' pattern ������ ���� '������;�����', ����. '���;�������[ $]';
                    '   ��������: '���' ���������, '�������[ $]' �����������.
                    ' ��������������, ��� ��� ��������� '������������'
                If Not patTest(comps(i), pats(0)) Then   ' 1-� ����� - false
                    If patTest(comps(i), pats(1)) Then   ' 2-� ����� �.�. true
                        patTest = True
                        GoTo Ex              ' ������ �������
                    End If
                    GoTo nextComp
                End If
nextComp:
            Next i
        Else
            pat = Replace(pat, "~", ",")
            .Pattern = pat
            patTest = .Test(longTxt)
    '        If .test(longTxt) Then
    '            patTest = "found: '" & pat & "' in: '" & longTxt & "'"
    '        Else
    '            patTest = "Not found: '" & pat & "' in: '" & longTxt & "'"
    '        End If
        End If
    End With
Ex:
End Function