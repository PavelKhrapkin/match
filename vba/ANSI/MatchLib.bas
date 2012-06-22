Attribute VB_Name = "MatchLib"
'---------------------------------------------------------------------------
' ���������� ����������� ������� MatchSF-1C
'
' �.�.�������, �.���� 19.6.2012
'
' - ModStart(SheetN, MSG)       - ������ ������ ������ � ������ SheetN
' - ModEnd(SheetN)              - ���������� ������, ����������� � ������ SheetN
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
' - TxDate(D)                   - �������������� ��������� ������ D � Date
' T testTxDate                  - ������� TxDate
' - DDMMYYYY(d)                 - �������������� ���� d � ��������� ������ DDMMYYYY
' - Dec(a)                      - ������ ����� � � ���� ������ � ���������� ������
' - EOL(SheetN)                 - ���������� ����� ��������� ������ ����� SheetN
' - ClearSheet(SheetN, HDR_Range) - ������� ����� SheetN � ������ � ���� �����
' - SheetSort(SheetN, Col)      - ���������� ����� SheetN �� ������� Col
' - SheetDedup(SheetN, Col)     - c��������� � ������������ SheetN �� ������� Col
' - SheetDedup2(SheetN, ColSort,ColAcc) - ���������� � ������� ����� SheetN
'                                 �� �������� ColSort, ColAcc
' - DateCol(SheetN, Col)        - �������������� ������� Col �� ������ � ����
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

Function ModStart(SheetN, Msg, Optional P As Boolean = False) As Integer
'
' ������������ - ������ ������ ������ � ������ SheetN, ���������� ���������� �����
'  23.3.2012
'  19.4.12 - ���������� ������ Excel
'  12.6.12 - Select Case � ����������� �� ��������� ����� SheetN

    Doing = Msg
    With Application
        .DisplayStatusBar = True
        .StatusBar = Msg
' ��� ��������� Excel ��������� ����� � ��.
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    ActiveSheet.DisplayPageBreaks = False
    Call AutoFilterReset(SheetN)
    
' ---- ����������� EOL ��� ���� �������� ������
    EOL_DogSheet = EOL(DOG_SHEET) - DOGRES
    EOL_SF = EOL(SF) - SFresLines
    EOL_SFD = EOL(SFD) - SFresLines
    EOL_SFopp = EOL(SFopp) - SFresLines
    EOL_SFacc = EOL(SFacc) - SFresLines
    EOL_Acc1C = EOL(Acc1C) - ACC1C_RES
    EOL_ADSKfrSF = EOL(ADSKfrSF) - SFresLines
    EOL_Stock = EOL(STOCK_SHEET)
    EOL_PaySheet = EOL(PAY_SHEET) - PAY_RESLINES
    EOL_SFlnkADSK = EOL(SF_PA) - SFresLines
    
    Select Case SheetN
    Case PAY_SHEET:     ModStart = EOL_PaySheet
    Case DOG_SHEET:     ModStart = EOL_DogSheet
    Case Acc1C:         ModStart = EOL_Acc1C
    Case STOCK_SHEET:   ModStart = EOL_Stock
    Case SF:            ModStart = EOL_SF
    Case SFD:           ModStart = EOL_SFD
    Case SFacc:         ModStart = EOL_SFacc
    Case SF_PA:         ModStart = EOL_SFlnkADSK
    Case Else:
        ModStart = EOL(SheetN)
    End Select
' ----
    ExRespond = True
    
    Range("A1:A" & ModStart).EntireRow.Hidden = False
    If P Then
        With ProgressForm
            .Show vbModeless
            .ProgressLabel.Caption = Doing
        End With
    End If
    LogWr ""
    LogWr (Doing)
End Function
Sub ModEnd(SheetN)
'
' ������������ ���������� ������ ������ � ������ SheetN
'  15.2.2012
'  19.4.12  - �������������� ������ Excel
'  19.6.12  - ������� ���� � ����� � ����� ������ ������� SheetN

    Dim i
    Dim Col As Integer  '= ����� ������� � SheetN

    Col = Sheets(SheetN).UsedRange.Columns.count
    Sheets(SheetN).Cells(1, Col + 1) = Now
    
    i = AutoFilterReset(SheetN)
    ActiveSheet.Range("A" & i).Select
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
Sub MS(Msg)
'
'   - MS(Msg)- ����� ��������� �� ����� � � LogWr
'   11.6.12
    ErrMsg TYPE_ERR, Msg, False
End Sub

Sub ErrMsg(ErrMode, Msg, Optional ByVal contRequest As Boolean = True)
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
        
        If contRequest Then
            Respond = MsgBox(Msg & vbCrLf & vbCrLf & "����������?", vbYesNo)
            If Respond = vbNo Then
                ExRespond = False
                Stop
            End If
        Else
            MsgBox Msg
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

    Dim N   ' ����� ������ � Log
    
    N = Sheets(Log).Cells(1, 4)
    N = N + 1
    Sheets(Log).Cells(N, 1) = Date
    Sheets(Log).Cells(N, 2) = Time
    Sheets(Log).Cells(N, 3) = Msg
    Sheets(Log).Cells(1, 4) = N
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
Function TxDate(D) As Date
'
' - TxDate(D)   - �������������� ��������� ������ D � Date
'
    If IsDate(D) Then
        TxDate = D
    Else
        TxDate = "1.1.1900"
    End If
End Function
Sub testTxDate()
    Dim A(4) As Date
    A(1) = TxDate("15/6")
    A(2) = TxDate("")
    A(3) = TxDate(Now)
End Sub
Function Dec(A) As String
'
' �������������� ����� � � ��������� ������ � ���������� ������
'   14.2.2012

    Dec = "'" & WorksheetFunction.Substitute(A, ",", ".")
End Function
Function EOL(SheetN)
'
' ���������� ���������� ����� � ����� SheetN � ������ ��������� ������ �����
'   20/1/2012
'   4/2/2012 - ��������� ������ On Error
'   20/2/2012 - ��������� Option Explicit
'   12.5.12 - Sheets(SheetN).Select ��������

    Dim i, Col
    
    On Error GoTo Err
    
    EOL = Sheets(SheetN).UsedRange.Rows.count
    Col = Sheets(SheetN).UsedRange.Columns.count
    Do
        For i = 1 To Col
            If Sheets(SheetN).Cells(EOL, i) <> "" Then Exit Do
        Next i
        If EOL <= 1 Then Exit Do
        EOL = EOL - 1       ' ������ UsedRange ��������� ������ ������,
    Loop                    '   .. ��������, ���� � ������ ���� ��������� ������
    Exit Function

Err: MsgBox "������ � ������ �� ����� " & SheetN & " � ������ (" & _
        i & "," & EOL & ")"
    Stop
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
'  12.6.12 - �������������� ����� �� ������ ������� �� ������ ������ HDR_Range

    Dim Col As Range    '= ������� ������� � ����� SheetN
    Dim W               '= ������ ������� �����
    
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
    
' -- ����������� ������� ����� �� ������, ��������� �� ������ ������ HDR_Range
    For Each Col In Sheets(SheetN).Columns
        W = Col.Cells(2, 1)
        If IsNumeric(W) And W > 0 And W < 200 Then Col.ColumnWidth = W
    Next Col
    
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
                Or smb = "�" Or smb = "�" _
                Or smb = "�") Then
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
