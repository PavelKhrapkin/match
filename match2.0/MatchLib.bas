Attribute VB_Name = "MatchLib"
'---------------------------------------------------------------------------
' ���������� ����������� ������� "match 2.0"
'
' �.�.�������, �.���� 31.7.2012
'
' - ModStart(Report)            - ������ ������ ������ � ������ SheetN
' - PublicVarInit()             - ������������� ���������� ���������� EOL � ��
' - ModEnd(SheetN)              - ���������� ������, ����������� � ������ SheetN
' - WrTOC()                     - ���������� Publoc RepTOC � TOCmatch
' - GetRep(RepName)             - ������� � ��������� ����� ������ RepName
' - GetMatch()                  - ��������� ���� match.xlsm �� �������� DirDBs
' - FileOpen(RepFile)           - ���������, ������ �� RepFile, ���� ��� - ���������
' - InsMyCol(F)                 - ��������� ������� � ���� ����� �� ������� � F
' - InsSummary(SheetN, F)       - ��������� ������ F � ��� EL �� ����� �����
' - MS(Msg)                     - ����� ��������� �� ����� � � LogWr
' - ErrMsg(ErrMode, MSG)        - ����� ��������� �� ������ � Log � �� �����
' - LogWr(msg)                  - ������ ��������� msg � Log list
'(*)LogReset()                  - ����� � ������� Log �����
' - ActiveFilterReset(SheetN)   - ����� � ����������� ����������� ����� SheetN
' - SheetsCtrlH(SheetN, FromStr, ToStr) - ������ ������ FromStr �� ToStr
'                                 � ����� SheetN
' - Pnt(SheetN,Col,Criteria,Color,Mode) - ������� �� SheetN ������ Color �� �������
' - PerCent(Row, Col)           - �������������� ������ (Row,Col) � ����������
' - CurCode(Row, Col, CurCol)   - ������ ������ (Row,Col) �� ���� ������ � CurCol
' - CurRate(Cur)                - ���������� ���� ������ � ����� �� ���� Cur ��� We
' - CurISO(Cur1C)               - ���������� ��� ������ � ��������� ISO
' - DDMMYYYY(d)                 - �������������� ���� d � ��������� ������ DDMMYYYY
' - Dec(a)                      - ������ ����� � � ���� ������ � ���������� ������
' - EOL(SheetN)                 - ���������� ����� ��������� ������ ����� SheetN
' - CSmatch(Val,Col)            - Case Sensitive match - ���������� ����� ������ � Val
'                                 � ������� Col. ���� Val �� ������- ���������� 0.
'                                 ���� ��� ������ Val ������ ���� Selected.
' - ClearSheet(SheetN, HDR_Range) - ������� ����� SheetN � ������ � ���� �����
' - SheetSort(SheetN, Col)      - ���������� ����� SheetN �� ������� Col
' - SheetDedup(SheetN, Col)     - c��������� � ������������ SheetN �� ������� Col
' - SheetDedup2(SheetN, ColSort,ColAcc) - ���������� � ������� ����� SheetN
'                                 �� �������� ColSort, ColAcc
' - DateCol(SheetN, Col)        - �������������� ������� Col �� ������ � ����
' - DateSort(SheetN, Col)       - �������������� ������� Col �� ���������� ������� � Date
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

Option Explicit

' ��������� - ������������ ������
    Public Const rgbADSK = 2162853  ' ����������� ������-���������� ���� ��� Autodesk
    Public Const LimeG = 3329330    ' RGB(50, 205, 50)  �������-�������
    Public Const Antique = 11583693 ' RGB(205, 192, 176) ����� - ����-����������
    Public Const Gray = 8750469     ' �����
    
    Public Const Log = "Log"        ' Log ����

Sub ModStart(Report)
'
' - ModStart(Report)    - ������ ������ � ������� Report, �������� � �������������
'
'  26.7.12  - ���������� ��� match 2.0

    GetRep TOC
    
    Select Case Report
    Case REP_1C_P_LOAD:
        Doing = "��������� ����� ����� �� �������� 1� � ���� 1C.xlsm"
        GetRep SF
        GetRep PAY_SHEET
''        CheckSheet PAY_SHEET
''        EOL_PaySheet = RepTOC.EOL
    Case REP_1C_P_PAINT:
        Doing = "������������ ���� �������� ���� 1C.xlsm"
    Case REP_1C_SFACCFIL:
        Doing = "���������� ������� 1 ��� ����� ��������"
        GetRep PAY_SHEET
        EOL_PaySheet = RepTOC.EOL
''''''''''''        EOL_SFacc = EOL(SFacc, F_SFDC) - SFresLines
    Case REP_SF_LOAD:
        Doing = "�������� �������� �� Salesforce - SF"
        Set DB_1C = Workbooks.Open(DirDBs & F_1C, UpdateLinks:=False, ReadOnly:=True)
        GetRep PAY_SHEET
        EOL_PaySheet = RepTOC.EOL
        GetRep SF
        EOL_SF = RepTOC.EOL
''        CheckSheet PAY_SHEET, 1, PAYDOC_COL, Stamp1Cpay1
''        CheckSheet PAY_SHEET, 1, PAYDATE_COL, Stamp1Cpay2
''        EOL_PaySheet = EOL(PAY_SHEET) - PAY_RESLINES
''        EOL_SFacc = EOL(SFacc, F_SFDC) - SFresLines
'''        P = True
'    Case REP_1C_�_LOAD:
'    Case Acc1C:
    Case Else:
        ErrMsg FATAL_ERR, "�������� ����������� �����"
        End
    End Select

'''''''    With Application
'''''''        .DisplayStatusBar = True
'''''''        .StatusBar = Msg
'''''''' ��� ��������� Excel ��������� ����� � ��.
'''''''        .ScreenUpdating = False
'''''''        .Calculation = xlCalculationManual
'''''''        .EnableEvents = False
'''''''        .DisplayAlerts = False
'''''''    End With
'''''''    ActiveSheet.DisplayPageBreaks = False
'''''''    Call AutoFilterReset(SheetN)
'''''''
'''''''' ---- ����������� EOL ��� ���� �������� ������
'''''''    EOL_DogSheet = EOL(DOG_SHEET) - DOGRES
'''''''    EOL_SF = EOL(SF) - SFresLines
'''''''    EOL_SFD = EOL(SFD) - SFresLines
'''''''    EOL_SFopp = EOL(SFopp) - SFresLines
'''''''    EOL_SFacc = EOL(SFacc) - SFresLines
'''''''    EOL_Acc1C = EOL(Acc1C) - ACC1C_RES
'''''''    EOL_ADSKfrSF = EOL(ADSKfrSF) - SFresLines
'''''''    EOL_Stock = EOL(STOCK_SHEET)
'''''''    EOL_PaySheet = EOL(PAY_SHEET) - PAY_RESLINES
'''''''    EOL_SFlnkADSK = EOL(SF_PA) - SFresLines
'''''''
'''''''    Select Case SheetN
'''''''    Case PAY_SHEET:     ModStart = EOL_PaySheet
'''''''    Case DOG_SHEET:     ModStart = EOL_DogSheet
'''''''    Case Acc1C:         ModStart = EOL_Acc1C
'''''''    Case STOCK_SHEET:   ModStart = EOL_Stock
'''''''    Case SF:            ModStart = EOL_SF
'''''''    Case SFD:           ModStart = EOL_SFD
'''''''    Case SFacc:         ModStart = EOL_SFacc
'''''''    Case SF_PA:         ModStart = EOL_SFlnkADSK
'''''''    Case Else:
'''''''        ModStart = EOL(SheetN)
'''''''    End Select
'''''''' ----
    ExRespond = True

'    Range("A1:A" & ModStart).EntireRow.Hidden = False
    With ProgressForm
        .Show vbModeless
        .ProgressLabel.Caption = Doing
    End With
    LogWr ""
    LogWr (Doing)
End Sub
Sub ModEnd()
'
' - ModEnd() - ������������ ���������� ������ ������
'  15.2.2012
'  19.4.12  - �������������� ������ Excel
'  2.7.12  - match 2.0
' 20.7.12 - ������������ TOCmatch to RepTOC

    WrTOC
    Close

'    i = AutoFilterReset(SheetN)
'    ActiveSheet.Range("A" & i).Select
    ProgressForm.Hide
' ��������������� ����� Excel � ��
    With Application
        .StatusBar = False
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayStatusBar = True
        .DisplayAlerts = True
    End With
    ActiveSheet.DisplayPageBreaks = True
    LogWr (Doing & " - ������!")
End Sub
Sub WrTOC()
'
' - WrTOC() - ���������� Public RepTOC � ���������� match.Sheets(1)
'   25.7.2012

    Dim i As Long
    
    For i = 4 To BIG
        If DB_MATCH.Sheets(1).Cells(i, TOC_REPNAME_COL) = RepTOC.SheetN Then GoTo FoundRep
    Next i
    GoTo FatalRepTOC

FoundRep:
    With DB_MATCH.Sheets(1)
        .Cells(i, TOC_DATE_COL) = RepTOC.Dat
        .Cells(i, TOC_REPNAME_COL) = RepTOC.Name
        .Cells(i, TOC_EOL_COL) = RepTOC.EOL
        .Cells(i, TOC_MYCOL_COL) = RepTOC.MyCol
        .Cells(i, TOC_RESLINES_COL) = RepTOC.ResLines
        .Cells(i, TOC_MADE_COL) = RepTOC.Made
        .Cells(i, TOC_NEXTREP_COL) = RepTOC.NextStep
        .Cells(i, TOC_REPFILE_COL) = RepTOC.RepFile
        .Cells(i, TOC_SHEETN_COL) = RepTOC.SheetN
        .Cells(i, TOC_STAMP_COL) = RepTOC.Stamp
        .Cells(i, TOC_STAMP_TYPE_COL) = RepTOC.StampType
        .Cells(i, TOC_STAMP_R_COL) = RepTOC.StampR
        .Cells(i, TOC_STAMP_C_COL) = RepTOC.StampC
        .Cells(i, TOC_CREATED_COL) = RepTOC.CreateDat
        .Cells(i, TOC_PARCHECK_COL) = RepTOC.ParChech
        .Cells(i, TOC_REPLOADER_COL) = RepTOC.Loader
    End With
    Exit Sub
    
FatalRepTOC:
    ErrMsg FATAL_ERR, "WrTOC: �������� ��������� Rep��� ��� ������ " & RepTOC.SheetN
    Stop
    End

End Sub
Function GetRep(RepName) As TOCmatch
'
' - GetRep(RepName) - ������� � ��������� ����� ������ RepName
'   26.7.12

    Dim i As Long
    
    If DB_MATCH Is Nothing Then
        Set DB_MATCH = FileOpen(F_MATCH)
        GetRep = GetRep(TOC)        ' ��� TOCmatch - ��������
            
        DirDBs = DB_MATCH.Path & "\"
        If DB_MATCH.Sheets(TOC).Cells(1, TOC_F_DIR_COL) <> DirDBs Then
            Dim Respond As Integer
            Respond = MsgBox("���� <match.xlsx> �������� �� ���������� �����!" _
                & vbCrLf & vbCrLf & "��� ������ ������� ������ DBs? ", vbYesNo)
            If Respond <> vbYes Then End
    '** ����� DirDBs ������� � TOCmatch � �� ��������������� ����
            DB_MATCH.Sheets(TOC).Cells(1, TOC_F_DIR_COL) = DirDBs
            Dim F_match_env As Workbook ' ��������������� ���� c DirDBs
            Set F_match_env = Workbooks.Open(F_match_environment)
                ' ��� ���� ��� ������ �� TOCmatch ������ ���� ��������!
            For i = 4 To EOL(TOC, DB_MATCH)
                GetRep DB_MATCH.Sheets(TOC).Cells(i, TOC_REPNAME_COL)
            Next i
            
            With F_match_env.Sheets(1)
                .Cells(1, 1) = Now
                .Cells(1, 2) = DirDBs
            End With
            F_match_env.Close
            Exit Function
        End If
    End If
    
    With DB_MATCH.Sheets(TOC)
        For i = 4 To EOL(TOC, DB_MATCH)
            If .Cells(i, TOC_REPNAME_COL) = RepName Then GoTo FoundRep
        Next i
        GoTo FatalRep

FoundRep:
        RepTOC.Dat = .Cells(i, TOC_DATE_COL)
        RepTOC.Name = .Cells(i, TOC_REPNAME_COL)
        RepTOC.EOL = .Cells(i, TOC_EOL_COL)
        RepTOC.MyCol = .Cells(i, TOC_MYCOL_COL)
        RepTOC.ResLines = .Cells(i, TOC_RESLINES_COL)
        RepTOC.Made = .Cells(i, TOC_MADE_COL)
        RepTOC.NextStep = .Cells(i, TOC_NEXTREP_COL)
        RepTOC.RepFile = .Cells(i, TOC_REPFILE_COL)
        RepTOC.SheetN = .Cells(i, TOC_SHEETN_COL)
        RepTOC.Stamp = .Cells(i, TOC_STAMP_COL)
        RepTOC.StampType = .Cells(i, TOC_STAMP_TYPE_COL)
        RepTOC.StampR = .Cells(i, TOC_STAMP_R_COL)
        RepTOC.StampC = .Cells(i, TOC_STAMP_C_COL)
        RepTOC.CreateDat = .Cells(i, TOC_CREATED_COL)
        RepTOC.ParChech = .Cells(i, TOC_PARCHECK_COL)
        RepTOC.Loader = .Cells(i, TOC_REPLOADER_COL)
    End With
    
'---- �������� ������ ----------
    Dim StR As Long, StC As Long
    Dim TestedStamp As String
    With RepTOC
        Select Case .RepFile
        Case F_MATCH:
            RepMatch = RepTOC
        Case F_1C:
            Set DB_1C = FileOpen(.RepFile)
            Rep1C = RepTOC
        Case F_SFDC:
            Set DB_SFDC = FileOpen(.RepFile)
            RepSF = RepTOC
        Case F_ADSK:
            Set DB_ADSK = FileOpen(.RepFile)
            RepADSK = RepTOC
        Case F_STOCK:
            Set DB_STOCK = FileOpen(.RepFile)
            RepStock = RepTOC
        Case Else: GoTo FatalRep
        End Select
            
        StR = .StampR
        If .RepFile = F_SFDC Then StR = StR + .EOL
        StC = .StampC
        If .Made <> REP_LOADED Then StC = StC + .MyCol
        TestedStamp = Workbooks(RepTOC.RepFile).Sheets(.SheetN).Cells(StR, StC)
        If .StampType = "=" Then
            If .Stamp <> TestedStamp Then GoTo FatalRep
        ElseIf .StampType = "I" Then
            If InStr(LCase$(TestedStamp), LCase$(.Stamp)) = 0 Then GoTo FatalRep
        Else
            ErrMsg FATAL_ERR, "���� � ��������� TOCmatch: ��� ������ =" & .StampType
        End If
    End With
    GetRep = RepTOC
    Exit Function
FatalRep:
    ErrMsg FATAL_ERR, "GetRep: ������ �� ������������� � ��� ������ " & RepName
    Stop
    End
End Function
Function FileOpen(RepFile) As Workbook
'
' - FileOpen(RepFile)   - ���������, ������ �� RepFile, ���� ��� - ���������
'   26.7.12
    
    Dim W As Workbook
    For Each W In Application.Workbooks
        If W.Name = RepFile Then
            W.Activate
            Set FileOpen = W
            Exit Function
        End If
    Next W
    
    If DirDBs = "" Then
        Dim F_match_env As Workbook ' ��������������� ���� c DirDBs
        Set F_match_env = Workbooks.Open(F_match_environment)
        DirDBs = F_match_env.Sheets(1).Cells(2, 1)
        F_match_env.Close
    End If
    
    Set FileOpen = Workbooks.Open(DirDBs & RepFile, UpdateLinks:=False)
End Function
Sub NextRep(RepName, MadeStep, NextStep)
'
' - NextRep(RepName, MadeStep, NextStep) - ���������� � ��� ��������� RepName
'   20.7.12

    RepTOC.Dat = Now
    RepTOC.Made = MadeStep
    RepTOC.NextStep = NextStep
    WrTOC
End Sub
Sub InsMyCol(F, FS)
'
' - InsMyCol(F) - ��������� ������� � ���� ����� �� ������� F � ����� �� FS
'  28.7.12
 
    Dim i As Integer
    If RepTOC.Made <> REP_LOADED Then Exit Sub
    
    With Workbooks(RepTOC.RepFile).Sheets(RepTOC.SheetN)
        .Activate
'---- ��������� ������� �� ����� MyCol
        For i = 1 To RepTOC.MyCol
            .Cells(1, 1).EntireColumn.Insert
        Next i
'---- ������ ������ ����������� �������
        For i = 1 To RepTOC.MyCol
            .Columns(i).ColumnWidth = Range(F).Cells(3, i)
        Next i
'---- �������� ������� MyCol �� ����� �� EOL
        Sheets("Forms").Range(F).Copy Destination:=.Cells(1, 1)
        .Range(.Cells(2, 1), .Cells(RepTOC.EOL, RepTOC.MyCol)).FillDown
    End With
'---- ��������� ����� �� FS
    Range(FS).Copy _
        Destination:=Workbooks(RepTOC.RepFile).Sheets(RepTOC.SheetN) _
        .Cells(RepTOC.EOL + RepTOC.ResLines - Range(F).Rows.count + 1, 1)
    
    RepTOC.Made = REP_INSMYCOL
    WrTOC
End Sub
Sub InsSummary(F)
'
' - InsSummary(F) - ��������� ������ (�����) �� ������� F
'  25.7.12
    
    If RepTOC.Made <> REP_INSMYCOL Then Exit Sub
    Workbooks(RepTOC.RepFile).Sheets(RepTOC.SheetN).Activate
    Range(F).Copy _
        Destination:=Workbooks(RepTOC.RepFile).Sheets(RepTOC.SheetN) _
        .Cells(RepTOC.EOL + RepTOC.ResLines - Range(F).Rows.count + 1, 1)
End Sub
Sub MS(Msg)
'
'   - MS(Msg)- ����� ��������� �� ����� � � LogWr
'   11.6.12
    ErrMsg TYPE_ERR, Msg
End Sub

Sub ErrMsg(ErrMode, Msg)
'
' - ErrMsg(ErrMode, MSG) - ����� ��������� �� ������ � Log � �� �����
'                          ���� ErrMode ���������� � Declaration
'   31.5.12

    Dim ErrType As String, Respond As String

    Select Case ErrMode
    Case WARNING:
        LogWr "< WARNING > " & Msg
        Exit Sub
        
    Case TYPE_ERR:
        LogWr "��������:" & Msg
        Respond = MsgBox(Msg & vbCrLf & vbCrLf & "����������?", vbYesNo)
        If Respond = vbNo Then
            ExRespond = False
            Stop
        End If
        Exit Sub
        
    Case FATAL_ERR:
Fatal:  ErrType = "<! ERROR !> "
        LogWr ErrType & Msg
        MsgBox Msg, , ErrType
'        Stop
        Exit Sub
    Case Else:
        ErrMode = FATAL_ERR
        GoTo Fatal
    End Select
End Sub
Sub LogWr(Msg)
'
' ������ ��������� msg � Log-����
'   15.2.2012
'   26.6.12 - match 2.0

    Dim N   ' ����� ������ � Log
    
    With ActiveWorkbook.Sheets(Log)
        N = .Cells(1, 4)
        N = N + 1
        .Cells(N, 1) = Date
        .Cells(N, 2) = Time
        .Cells(N, 3) = Msg
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
' ������������ ������ � ������ ������� ����� SheetN � ������ ������
'      ���������� ���������� ����� � SheetN
' 16.1.2012

    Sheets(SheetN).Select
    ActiveSheet.AutoFilterMode = False  ' ���������� ����� �������
    ActiveWindow.FreezePanes = False    ' Top Row Freeze
    Rows("1:1").AutoFilter              ' ��������/��������� AutoFilter
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    AutoFilterReset = Sheets(SheetN).UsedRange.Rows.count
    Range("A" & AutoFilterReset).Activate ' �������� ������ ����� �����
End Function
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
Sub Pnt(Col, Criteria, Color, Optional Mode As Integer = 0)
'
' ������������ �������� ������� Col �� �������� Criteria � ���������� � Color
' ���� Mode = 0 ��� �� ������ - ���������� ���� ���, ����� ������ Col
'   26.1.2011

    AllCol = ActiveSheet.UsedRange.Columns.count
    Range(Cells(1, 1), Cells(Lines, AllCol)).AutoFilter _
                            Field:=Col, Criteria1:=Criteria
    If Mode = 0 Then
        Range(Cells(2, 2), Cells(Lines, AllCol)).Interior.Color = Color
    Else
        Range(Cells(2, Col), Cells(Lines, Col)).Interior.Color = Color
    End If
    If Criteria = "�� ���������" Then   ' "�� ���������" - �������������
        Range(Cells(2, 2), Cells(Lines, AllCol)).Font.Strikethrough = True
    End If
    ActiveSheet.UsedRange.AutoFilter Field:=Col
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

    Dim S

    CurRate = 1
    If Cur = "RUB" Or Trim(Cur) = "" Then Exit Function
    S = WorksheetFunction.VLookup(Cur, Sheets(We).Range("RUB_Rate"), 2, False)
    CurRate = Replace(S, ".", ",")
End Function
Function CurISO(Cur1C)
'
' ���������� ��� ������ � ��������� ISO, ������������ ��� �� ���� 1�
'   18.3.2012

    CurISO = ""
    On Error Resume Next
    CurISO = WorksheetFunction.VLookup(Cur1C, Range("Currency"), 2, False)
    On Error GoTo 0
End Function
Function DDMMYYYY(D) As String
'
' �������������� ���� d � ��������� ������ DDMMYYYY
'   14.2.2012
    DDMMYYYY = Day(D) & "." & Month(D) & "." & Year(D)
End Function
Function Dec(A) As String
'
' �������������� ����� � � ��������� ������ � ���������� ������
'   14.2.2012

    Dec = "'" & WorksheetFunction.Substitute(A, ",", ".")
End Function
Sub testEOL()
    Dim A, B, C
    A = EOL(1)
    B = EOL(2)
        Dim F As Workbook
        Set F = Workbooks.Open(F_SFDC, UpdateLinks:=True, ReadOnly:=True)
    C = EOL(1, F)
        F.Close SaveChanges:=False
End Sub
Function EOL(SheetN, Optional F = Null)
'
' - EOL(SheetN,[F]) - ���������� ���������� ����� � ����� SheetN ����� F
'   20/1/2012
'   4/2/2012 - ��������� ������ On Error
'   20/2/2012 - ��������� Option Explicit
'   12.5.12 - Sheets(SheetN).Select ��������
'   24.6.12 - AllCol - Public
'   29.6.12 - match 2.0 - ������� ���� S, ���� ������ ��������

    Dim i

    If IsNull(F) Then
        Set F = ThisWorkbook
    End If
    
    With F.Sheets(SheetN)
        EOL = .UsedRange.Rows.count
        AllCol = .UsedRange.Columns.count
        Do
            For i = 1 To AllCol
                If .Cells(EOL, i) <> "" Then Exit Do
            Next i
            If EOL <= 1 Then Exit Do
            EOL = EOL - 1       ' ������ UsedRange ��������� ������ ������,
        Loop                    '   .. ��������, ���� � ������ ���� ��������� ������
    End With
    Exit Function

Err: MsgBox "������ � ������ �� ����� " & SheetN & " � ������ (" & _
        i & "," & EOL & ")"
    Stop
End Function
Function CSmatch(Val, Col) As Double
'
' - CSmatch(Val,Col) - Case Sensitive match ���������� ����� ������ � Val � ������� Col.
'                   ���� Val �� ������- ���������� 0. ���� ��� ������ Val ������ ���� Selected.
' 8/7/12

'    Const BIG = 77777
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
Sub ClearSheet(SheetN, HDR_Range As Range)
'
' ������ ������� SheetN � ������� � ���� ��������� �� ����� �eader.HDR_Range
'   4.2.2012
'  11.2.2012 - ��������� ������������
'  10.3.12 - ��������� ������������ - �������� HRD_Range
'  25.3.12 - ����� C_Contr � C_ContrLnk
'  17.4.12 - ���� A_Acc - ����� �����������
'  18.4.12 - ���� A_Dic - ������� �����������
'  28.4.12 - ���� NewOrderList - ���� ����� �������
'  13.5.12 - ���� P_ADSKlink - ����� ������ ������ - ADSK
'  15.5.12 - ���� SF_PA ������ �������� � ����������� ADSK
'   6.6.12 - Delete ������ ����, ������� �����
'  11.6.12 - ����� A_Acc � AccntUpd
'  12.6.12 - ���� BTO_SHEET - ��� ��� ����� ���

' -- ������� ������ ����
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(SheetN).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
' -- ������� ����� ����
    Sheets.add After:=Sheets(Sheets.count)  ' ������� ����� ���� � ����� ������
    ActiveSheet.Name = SheetN
    ActiveSheet.Tab.Color = RGB(50, 153, 204)   ' Tab �������
   
    HDR_Range.Copy Sheets(SheetN).Cells(1, 1)   ' �������� ����� �� Header
    
    Select Case SheetN
    Case O_NewOpp:      EOL_NewOpp = 1
    Case P_Paid:        EOL_NewPay = 1
    Case C_Contr:       EOL_NewContr = 1
    Case C_ContrLnk:    EOL_ContrLnk = 1
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
Sub SheetSort(SheetN, Col)
'
' ��������� ���� SheetN �� ������� Col
'   22.1.2012
'   21.2.2012 - Option Explicit
'   19.4.12 - AutoFilterReset

    Dim Name As String

'    Sheets(SheetN).Select
    Call AutoFilterReset(SheetN)

    Name = ActiveSheet.Name
    
    With ActiveWorkbook.Worksheets(Name).AutoFilter.Sort
        .SortFields.Clear
        .SortFields.add key:=Cells(1, Col), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
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
Sub SheetDedup2(SheetN, ColSort, �olAcc)
'
' - SheetDedup2(SheetN, ColSort, ColAcc)  - ��������� ���� SheetN
'               �� ������� SortCol ����� �������� ColAcc ������ � "+"
'   23.5.2012

    Dim i As Integer, EOL_SheetN As Integer
    Dim prev As String, x As String
    Dim PrevAcc As String, NewAcc As String
    
    Call SheetSort(SheetN, ColSort)
    EOL_SheetN = EOL(SheetN)
    
    prev = "": i = 2
    With Sheets(SheetN)
        Do
            x = .Cells(i, ColSort)
            If x = prev Then
                PrevAcc = .Cells(i - 1, �olAcc)
                NewAcc = .Cells(i, �olAcc)
                If PrevAcc <> "" And NewAcc <> "" And PrevAcc <> NewAcc Then
                    PrevAcc = PrevAcc & "+" & NewAcc
                ElseIf PrevAcc = "" And NewAcc <> "" Then
                    PrevAcc = NewAcc
'                ElseIf PrevAcc <> "" And NewAcc = "" Then
'                ElseIf PrevAcc = "" And NewAcc = "" Then
'                   � ���� ��������� ������� ������ �� ������
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
Sub DateCol(SheetN, Col)
'
' �������������� ������� Col � ����� SheetN �� ������ ���� DD.MM.YY � ������ Date
'   20.4.12

    Dim i, DD, MM, yy As Integer
    Dim Dat As Date
    Dim D() As String
    
    For i = 1 To EOL(SheetN)
        D = split(Sheets(SheetN).Cells(i, Col), ".")
        If UBound(D) = 2 Then
            DD = D(0)
            If DD < 1 Or DD > 31 Then GoTo NXT
            MM = D(1)
            If MM < 1 Or MM > 12 Then GoTo NXT
            yy = D(2)
            Dat = DD & "." & MM & "." & yy
            Sheets(SheetN).Cells(i, Col) = Dat
        End If
NXT:
    Next i
End Sub
Sub DateSort(SheetN, Col)
'
' - DateSort(SheetN, Col) - �������������� ������� Col �� ���������� ������� � Date
'                           � ���������� �� ���� ������� �� ������ � ����� �����
'   31.7.12

    DateCol SheetN, Col
    SheetSortSheetN , Col
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
        
    With ProgressForm
        .ProgressFrame.Caption = Format(Pct, "0%")
        .LabelProgress.Width = Pct * .ProgressFrame.Width
        .Repaint
    End With
    
    Static t
    Dim R As String
    If t = 0 Then t = Timer
    If Timer - t > 20 Then
        t = Timer
        R = MsgBox("������?", vbYesNo)
        If R = vbNo Then ExRespond = False
    End If
    
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
    Dim A
    A = FindInLst("Autodesk Plant Design Suite Premium 2012 New SLM", "DIC_GoodADSK")
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
    Dim A
    A = FindInDIC("��������", "Goods")
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
    x = split(DicList, ",")
    
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
    Dim A As Boolean
    A = IsMatchList("", "�����,�����,���")
    A = IsMatchList("������", "�����,�����,���,���")
    A = IsMatchList("������", "�����,�����,���")
End Sub
