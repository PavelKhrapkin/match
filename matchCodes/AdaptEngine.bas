Attribute VB_Name = "AdaptEngine"
'---------------------------------------------------------------------------------------
' AdaptEngine - ��������� ���������.
'       * ������� - �������������� �������, ��������������� ��� ��������� ���� ���������
'       * �������� ��������� �������� � �������, � ��� �����:
'           1.(���.1) �������� �������������� �������� - �� ���������� ActiveSheet
'           2.(���.2) ��� ������� (������������ � �������� WP)
'               2.1 ������ (New, ->, Stop)
'               2.2 iLine - ������ �� ������ ����� iLine, ���������� ��� �������� WP_Adapt
'               2.3 ������ - �������� ��������������� ��� ������� ������ ����������
'               2.4 Select - �������� ������� OppSelect, ��������� ��������� �����-�������
'       * ������ ������ ������� �������� "�����" - ��������� �������
'       * ���, ���������� �������� �������, ������ ���� �� ������� ��������� ���������
'       * ������ ������ - Value - �������� ������� � ���������� Y, ����������� ����������
'           - ������ Select ������� ������, ��������� ��������� OppFilter ���� �������
'           - ����� ������, ��������� � Select ���� �������� ������ - �������� �� �������
'           - ��������� ���� ������� � ������ Value �������� ����������� ��� ������ SF
'       * ������ ������ - Width - ������ ��������� ������� � ������ ��������� ��������
'           - ����� - ������ � �������� �������� - ����� 1 ��
'           - �������� Dbl,..
'           - ��� ������� Select ������ ������� ������������ ������ ������ ������� Width
'       * ��������� ������ - Columns ��� � - ����������� ��������� ��������
'           - ������ ���������� ���������� � X_Parse
'           - #6 - ��������� � ������� 6 � Value ������ �������, � �� ActiveSheet
'           - !6 - ���� ��������� � �������, �� � Select � ������� �� ��������� ������
'       * ����� ������ - ������� - ���������� ���������� Y = Adapter (X)
'       * ������ ������ - Fetch - ����� ���������� � �������� ���������� ��� ����������
'         �� ������ ���������� ���������� ���� X = SFD/18:2, �� ���� X �� ��������� ������
'         ������������ ��� Lookup � ��������� SFD: ��� �������� ��������� � ������ 18, �
'         �������� � ������� 2 ��������� ������ ���������� �������� ��� ������� ��������.
'       * ��� ������� ��������������� � ������� Range � �������� �������, ��� �������������
'         � ������� Pass0 �� ������ ��������� �������. ��� ��������������� ������� �����
'         ��� ������_Pass0
'
' 21.10.13 �.�.�������, �.����
'   ������� ������:
' 11.11.12 - ��������� AdaptEngine �� ProcessEngine
'  7.12.12 - ������� ������� ������ "Dbl", "Txt", "Date" � ������ "width" � sub WP_Adapt
'  8.12.12 - ������ ��o��������� ������ � ������ width
' 14.12.12 - ��������� ��������� ������� � ������ PTRN_WIDTH (WrNewSheet)
' 17.12.12 - �������� ���� ������ ������� � testfmtCell()
' 19.12.12 - ������� ����������� ����� � Dbl � testfmtCell()
' 3.1.13   - ������� �������������� Adapt
' 5.1.13   - Merged with Google source
' 8.4.13   - ����� ������ #2 � ������ �������� "Columns" ������ "������� 2 ������� ������
'                                                                ��������� �����"
' 1.5.13   - �������� ������������� ������� ���������� � WrNewSheet
'
' - WrNewSheet(SheetNew, SheetDB, DB_Line[,IdOpp]) - ���������� ����� ������
'                               � ���� SheetNew �� ������ DB_Line ����� SheetDB
'...........................
' - WP_Adapt(F, iLine)    - ������� �� ����� WP �� ������� � iLine, ������������� ������,
'                           ��� F - �������� ��������� ����� ��������� (WP_Prototype)
' * WP_Adapt_Continue(Button) - ����������� ������ WP_Adapt ����� ������� ������ Button
'...........................
' S Adapt(F) - ��������� �������� �� ������� F
' - Adater(Request, X, F_rqst, IsErr) - ������������ X � �������� "Request"
'        � �������� ������� � ��������� F_rqst. IsErr=True - ������ � ��������
' - X_Parse(iRow, iCol, PutToRow, PutToCol)    - ������ ������ � - ��������� ��������
' - FetchDoc(F_rqst, X, IsErr) - ���������� ������ �� ���������� ���������
'                   �� ������� F_rqst ��� �������� ���� X. IsErr=True - ������

Option Explicit

'========== ��������� � ����� ���������� ��������� ==================
Const WP_CONTEXT_LINE = 8, WP_CONTEXT_COL = 4   ' ������ �������� iLine
Const WP_PAYMENT_LINE = 8                       ' ������ ������� � WP

Const EXT_PAR = "ExtPar"    ' ����� � ������� - ������� �������� ��������� �
Const EXT_PAR_2 = "ExtPar2" ' ����� � ������� - ������� ������� ��������� �
Const EXT_PAR_3 = "ExtPar3" ' ����� � ������� - ������� ������� ��������� �

Const PTRN_HDR = 1   ' �������� ������ - �������� ����� � �������
Const PTRN_VALUE = 2 ' �������� ������ - �������� - Value � �������
Const PTRN_WIDTH = 3 ' �������� ������ - ������ ������� � �������
Const PTRN_COLS = 4  ' �������� ������ ������� ������� � �������
Const PTRN_ADAPT = 5 ' �������� ������ ������ ��������� � �������
Const PTRN_FETCH = 6 ' �������� ������ ������ Fetch - ���������� �� ���-� � �������
Const PTRN_LNS = 6   ' ���-�� ����� � ������� �� ������ ������ ����� �� ������

Const FOR_PROCESS = "ForProcess"    '������� � ������� ��� ��������
                                    '.. � �������� ����� ��������
Const PTRN_SELECT = "Select"
Const OPP_BALKY = "��������� ��������� � ���"

Const WP_PROTOTYPE = "WP_Prototype"

Sub WrNewSheet(SheetNew As String, SheetDB As String, DB_Line As Long, _
    Optional ExtPar As String, Optional ExtPar2 As String, Optional ExtPar3 As String)
'
' - WrNewSheet(SheetNew, SheetDB, DB_Line[,ExtPar,ExtPar2]) - ����������
'           ����� ������ � ���� SheetNew �� ������ DB_Line ����� SheetDB
'   * ������������ ������ � ���������� "HDR_" & SheetNew.
'     ������� ������ ��������� � Headers ��� WP_Prototype
'   * ��������� � �������� ����� ��� <�����������>/<���1>,<���2>...
'   * ���� � ������� � ������ PTRN_COLS ������� "ExtPar", ���������� �������
'                                              �������� ExtPar = IdOpp
' 6.9.2012
' 26.10.12 - ��������� "�������" ������ � DB_TMP
' 27.10.12 - ������������� TOCmatch ��� "�������" ������
' 28.10.12 - �������� SheetDB - ���������� � ���� String
' 14.12.12 - ��������� ��������� ������� � ������ PTRN_WIDTH
' 28.01.13 - width � setColWidth ������ ������: ������/������
' 8.4.13   - ����� ������ #2 � ������ �������� "Columns" ������
'            "������� 2 ������� ������ ��������� �����"
' 1.5.13   - �������� ������� ���.���������� ��� Columns � ���� ExtArr/3,
'            �� ���� ������ ������� ����������� ������� ���.����������
' 28.08.13 - WrTOC SheetNew
'  6.09.13 - ��������������� ��������� ��� IsErr=True
'  8.09.13 - ������������� X_Pars
' 27.09.13 - ��������� ExtPar2 � ExtPar3

''    Dim Rnew As TOCmatch, Rdoc As TOCmatch
    Dim P As Range
    Dim i As Long
    Dim X As String         '= �������������� �������� � SheetDB
''    Dim sX As String        '���� � ������ PTRN_COLS �������
''    Dim sXarr() As String   '����� �������� � ������� ExtPar, ����., ExtPar/2
    Dim Y As String         '= ��������� ������ ��������
    Dim IsErr As Boolean    '=True ���� ������� ��������� ������
    Dim NewEOL As Long      '=EOL(SheetNew)
    Dim Width() As String
    
    GetRep (SheetNew)
''    Rnew = GetRep(SheetNew)
''    Rnew.EOL = EOL(Rnew.SheetN, DB_TMP) + 1
''    Rnew.Made = "WrNewSheet"
''    Rdoc = GetRep(SheetDB)
      
    With DB_TMP.Sheets(SheetNew)
        Set P = DB_MATCH.Sheets(Header).Range("HDR_" & SheetNew)
        NewEOL = EOL(SheetNew, DB_TMP) + 1
        For i = 1 To P.Columns.Count
            Width = Split(P.Cells(PTRN_WIDTH, i), "/")
            '------------------------
            X = X_Parse(SheetDB, SheetNew, P.Cells(PTRN_COLS, i), DB_Line, NewEOL, _
                    ExtPar, ExtPar2, ExtPar3)
            '--------------------------
                
            Y = Adapter(P.Cells(PTRN_ADAPT, i), X, P.Cells(PTRN_FETCH, i), IsErr, , DB_Line)
            
            If IsErr Then
                MS "WrNewSheet: ������ ��� ������ � ���� '" & SheetNew & "'" _
                    & vbCrLf & " �� ����� '" & SheetDB & "' ������=" & DB_Line _
                    & vbCrLf & " � �������=" & i & " =" & P.Cells(PTRN_COLS, i) _
                    & " (" & P.Cells(PTRN_HDR, i) & ")  X=" & X _
                    & vbCrLf & " �������=" & P.Cells(PTRN_ADAPT, i) _
                    & " Fetch =" & P.Cells(PTRN_FETCH, i)
                .Rows(NewEOL).Delete
                Exit For
            Else
                '-- ���������� � SheetNew �������� Y � ���������� ������� ������
                fmtCell DB_TMP, SheetNew, Width, Y, NewEOL, i
            End If
        Next i
    End With
    If Not IsErr Then
'!        RepTOC = Rnew
        WrTOC SheetNew
    End If
End Sub
Sub testWP_Adapt()
    WP_Adapt "HDR_WP", 2
    Stop
End Sub
Sub WP_Adapt(ByVal F As String, ByVal iPayLine As Long)
'
' - WP_Adapt(F, iPayLine)    - ������� �� ����� WP �� ������� � iPayLine, ������������� ������,
'                           ��� F - �������� ��������� ����� ��������� (WP_Prototype)
'
' 18.10.13
    
    Dim LocalTOC As TOCmatch
    Dim nOpp() As Long, iOpp As Long, i As Long
    
    StepIn
    
'---- ������� ������ ���� WP, �������� �����-������ �� WP_Prototype
    
    LocalTOC = GetRep(WP)

    With Workbooks(LocalTOC.RepFile)
        Application.DisplayAlerts = False
            On Error Resume Next
                .Sheets(WP).Delete
            On Error GoTo 0
        Application.DisplayAlerts = True
        DB_MATCH.Sheets(WP_PROTOTYPE).Copy Before:=.Sheets(1)
        .Sheets(1).Name = WP
        With .Sheets(LocalTOC.SheetN)
            .Tab.Color = rgbCoral
            .Rows("1:" & EOL(WP, DB_TMP)).Delete    ' ������� ������ �� �������
            
            Dim FF As Range:  Set FF = DB_MATCH.Sheets(WP_PROTOTYPE).Range(F)
            FF.Copy .Cells(1, 1)
    '---- ������ ������ � ��������� ����������� �������
            For i = 1 To FF.Columns.Count
''                If Not TraceWidth Then setColWidth DB_TMP.Name, WP, i, FF.Cells(3, i)
            Next i
    
            Strip 1, 2, 0           ' ������
            Strip 7, 8, iPayLine    ' ������
            Strip 13, 14, 0         ' �����������
            If .Cells(8, 7) = "" Then
                .Rows("19:24").Hidden = True    ' ���� �������� ���
            Else
                Strip 19, 20, 0     ' �������
            End If
        '-------------              ' �������
            nOpp = OppSelect(iPayLine)
            For iOpp = 1 To nOpp(0)
                Strip 25, 30 + iOpp, nOpp(iOpp)
            Next iOpp
           
            .Rows(26).Hidden = True
            If nOpp(0) = 0 Then
                .Rows("25:30").Hidden = True
                .Cells(32, 11) = _
                    "� Salesforce ��� ���������� ��������. " _
                    & "������� ������� ���� �� ������ [NewOpp], [->] ��� [STOP]"
                .Rows(32).Interior.Color = rgbRed
            End If
            .Cells(1, 5) = "'" & DirDBs & F_MATCH & "'!WP_Adapt_Continue"
            .Cells(WP_CONTEXT_LINE, WP_CONTEXT_COL) = iPayLine
            .Activate
        End With
    End With
'''''''''''''''''''''''''''''''''''
    End '''  ��������� VBA ''''''''
'''''''''''''''''''''''''''''''''''
End Sub
Sub Strip(ByVal iPattern As Long, ByVal iOut As Long, ByVal iLine As Long)
'
' - Strip(iPattern, iOut, iLine)   - ����� � ���� WP �� �������
'                       �� ������ iPattern � ������� � ������ iOut
'                       �� ������ �������� ��������� ����� iLine
' 21.10.13
    
    Dim LocalTOC As TOCmatch, IsErr As Boolean, Width() As String
    Dim X As String, Y As String, Rqst As String, X_rqst As String, F_rqst As String
    Dim iCol As Long
    
    With DB_TMP.Sheets(WP)
        If .Cells(iPattern, 1) <> "" Then
            LocalTOC = GetRep(.Cells(iPattern, 1)) ' ��������� ��������, � ������� �������� ������
        End If
        
        For iCol = 5 To .UsedRange.Columns.Count
            X = X_Parse(InDoc:=.Cells(iPattern, 1), _
                    OutDoc:=WP, _
                    X_rqst:=.Cells(iPattern - 1 + PTRN_COLS, iCol), _
                    iLine:=iLine, _
                    PutToRow:=iOut)

            Rqst = .Cells(iPattern - 1 + PTRN_ADAPT, iCol)
            F_rqst = .Cells(iPattern - 1 + PTRN_FETCH, iCol)
            
            Y = Adapter(Rqst, X, F_rqst, IsErr, iPattern - 1 + PTRN_VALUE, iLine, iCol, iOut)
            
            X = .Cells(iPattern + PTRN_COLS - 1, iCol)
            If X = "-1" Then Exit For
            If Not IsErr And X <> "" Then
                Width = Split(.Cells(iPattern + PTRN_WIDTH - 1, iCol), "/")
                fmtCell DB_TMP, WP, Width, Y, iOut, iCol
            End If
        Next iCol
        .Rows(iPattern - 1 + PTRN_COLS).Hidden = True
        .Rows(iPattern - 1 + PTRN_ADAPT).Hidden = True
        .Rows(iPattern - 1 + PTRN_WIDTH).Hidden = True
        .Rows(iPattern - 1 + PTRN_FETCH).Hidden = True
    End With
End Sub
Sub WP_Adapt_Continue(Button As String, iRow As Long)
'
' * WP_Adapt_Continue(Button) - ����������� ������ WP_Adapt ����� ������� ������ Button
'                             ���� ���������� ���������� �� WP_Select_Button.
' 8/10/12
' 20.10.12 - ��������� ������ "�������"
' 10.11.12 - bug fix - ����������� ����� WP � ������������ Namer Range
' 24.09.13 - WrNewSheet NewOpp � CloseDate
'  1.10.13 - SheetDedup
' 21.10.13 - ������� ������ "�������" ��� "�������"

    Dim Proc As String, Step As String, iStep As Long
    Dim iPayment As Long, OppId As String, IsErr As Boolean
    Dim AccId As String, DefName As String
    Dim Respond As String, PaymentGoodType As String, PaymentContract As String
'---- ��������� ��������� �� ����� WP, �� ���� ������ �������, ������� -----
    With ActiveSheet
        iPayment = .Cells(WP_CONTEXT_LINE, WP_CONTEXT_COL)
        OppId = .Cells(iRow, 6)
        AccId = .Cells(WP_CONTEXT_LINE, 5)
        PaymentGoodType = .Cells(WP_CONTEXT_LINE, 11)
        PaymentContract = ContrCod(.Cells(WP_CONTEXT_LINE, 21), .Cells(WP_CONTEXT_LINE, 22))
    End With
    
    If DB_TMP Is Nothing Then Set DB_TMP = FileOpen(F_TMP)
    If DB_1C Is Nothing Then Set DB_1C = FileOpen(F_1C)
    If DB_MATCH Is Nothing Then Set DB_MATCH = FileOpen(F_MATCH)
    
    With DB_MATCH.Sheets(Process)
''        .Activate
        Proc = .Cells(1, PROCESS_NAME_COL)
        Step = .Cells(1, STEP_NAME_COL)
        iStep = ToStep(Proc, Step)
''        .Cells(iStep, PROC_PAR2_COL) = iPayment + 1
    End With
    
    Select Case Button
    Case "STOP":
        StepOut Step, iStep
        
        SheetDedup NEW_CONTRACT, 1
        SheetDedup DOG_UPDATE, 1
        SheetDedup NEW_OPP, NEWOPP_OPPNAME_COL

        ProcStart Proc
        End
    Case "->":
    Case "NewOpp":
        Const FETCH_SFACC = "SFacc/" & SFACC_IDACC_COL & ":" & SFACC_ACCNAME_COL
        DefName = FetchDoc(FETCH_SFACC, AccId, IsErr) & "-" & PaymentGoodType _
            & " ������� " & PaymentContract
        Respond = MsgBox(DefName, vbYesNo, "��� ������ �������")
        If Respond = vbYes Then
            Dim CloseDate As String, PayVal As String
            CloseDate = DB_1C.Sheets(PAY_SHEET).Cells(iPayment, PAYDATE_COL)
            PayVal = DB_1C.Sheets(PAY_SHEET).Cells(iPayment, PAYRUB_COL)
            WrNewSheet NEW_OPP, PAY_SHEET, iPayment, DefName, CloseDate, PayVal
        End If
    Case "NewAcc":
    ' ���� �� ��������
'-------- ��������� ������ �� ������� ����� ��������
    Case "������� ->":
        WrNewSheet NEW_PAYMENT, PAY_SHEET, iPayment, OppId
    Case "�������  ->"
        WrNewSheet DOG_UPDATE, PAY_SHEET, iPayment, OppId
    Case Else
        ErrMsg FATAL_ERR, "WP_Adapt_Continue: - ����������� ������!"
    End Select
    
NextWP:
    Paid1C iPayment + 1
End Sub
Sub Adapt(Optional FromDoc As String = "", Optional ToDoc As String = "")
'
' S Adapt([FromDoc],[ToDoc]) - ��������� �������� �� �������, ����������� ������
'                              �� ���������. �������� - ActiveSheet - ������ �
'                              ������ ���������� ����.
'                   ������������ FromDoc � ToDoc - ������� � �������� ���������.
'                   �� ��������� ������ � ������ ���������� � �������� ��������.
'
' ����� F ����� ���:
'   �����   - ��������� �������. ����� ������������ � ������������� ����� InsMyCol
'   MyCol   - ������� � ��������� ����� �� ����� �� �����. ���� "V" - ������ �����
'   Width   - ������ �������
'   Columns - ����� ������� � �������� �����- ����� � ������ ���������� � ��������
'       - ������ ���� Columns - ������� ���� �������� ��� ���������
'       <0  - ����� �� ����� �� ��������
'   ������� - ������- ����� ��������, ��������������� � = <�������� �� Columns>
'   Fetch   - ������ �������������� ���������� ��� �������� �� ������ ����������
'
' 12.9.12
' 14.9.12 - ���� ������� �� ����� �������� - ��������� �������� �� ���������
' 26.9.12 - ��������� ������ � ������������� �������� Columns
'  3.1.13 - ������� ��������������
'  6.1.13 - Optional FromDoc � ToDoc - �� ��������� ActiveSheet
' 10.1.13 - ������� ToDoc - ������� ������ � ����� ����
' 24.1.13 - ����� fmyCell ��� ������ Y ������ � �������� ������
'  2.5.13 - � ������ ������� Column ������ ��������� ������ �� �������� ���� #n
' 24.8.13 - �������� ��������- ��������, � ��� ������ - ������� � TOC
' 31.8.13 - ��� ��������� ������ ������� �������, ���� �������� ��� � ForProcess
'  1.8.13 - ���� ThisProc - �� ������������ ������ ����� �� ������� <ThisProcess>
' 23.9.13 - �������������� ������ ����������� � AdaptPass0

    StepIn
    
    Dim FF As Range     '= ����� F
    Dim R As TOCmatch
    Dim Rqst As String, F_rqst As String, IsErr As Boolean
    Dim X As String, Y As String
    Dim i As Long, Col As Long, iX As Long, iTo As Long, sX As String
    
    Dim R_From As TOCmatch, R_To As TOCmatch
    Dim Width() As String
    Dim Proc As String  ' ���� � ������� ��� �������� ����� ��������
    Dim IsThisProc As Boolean, ThisProcCol As Long, iProc As Long, sProc() As String
    Dim Pass0Arr() As Long, iPass As Long, iPassArr As Long
    
    ' ��������������
    Dim tot1 As Single, beg1 As Single: tot1 = 0
    Dim tot2(40) As Single, beg2(40) As Single
    Dim profileStr As String
    
    R = GetRep(ActiveSheet.Name)
    
    Set FF = DB_MATCH.Sheets(Header).Range( _
        DB_MATCH.Sheets(TOC).Cells(R.iTOC, TOC_FORMNAME))
        
    Pass0Arr = AdaptPass0(DB_MATCH.Sheets(TOC).Cells(R.iTOC, TOC_FORMNAME))
   
    If FromDoc = "" Then
        R_From = R
    Else
        R_From = GetRep(FromDoc)
    End If
    If ToDoc = "" Then
        R_To = R
        iTo = 2
    Else
        R_To = GetRep(ToDoc)
        NewSheet ToDoc
        iTo = R_To.EOL
    End If
    Proc = "": ThisProcCol = 0: IsThisProc = False  '=���������� ��� ��������� �������� �� �������
    
    beg1 = Timer()                  ' ��������������
    For i = 2 To R_From.EOL
        Progress i / R_From.EOL
        iTo = iTo + 1
        If ToDoc = "" Then iTo = i
        For iPass = 1 To 2
            For Col = 1 To FF.Columns.Count
                beg2(Col) = Timer()       ' ��������������
                                ' ���� ThisProc - ���������� ����� ����� �������
                If IsThisProc And Col < ThisProcCol Then Col = ThisProcCol
                If iPass = 1 Then
                    For iPassArr = 1 To Pass0Arr(0)
                        If Pass0Arr(iPassArr) = Col Then
                            GoTo HandlAdapt
                        Else
                            GoTo NextCol
                        End If
                    Next iPassArr
                End If
HandlAdapt:
             '--- ���������� X - ��������� ��������
                sX = FF(PTRN_COLS, Col)
                If sX = "" Then GoTo NextCol
                If IsNumeric(sX) Then
                    iX = sX
                    If iX < 0 Then Exit For
                    X = Workbooks(R_From.RepFile).Sheets(R_From.SheetN).Cells(i, iX)
                ElseIf Left(sX, 1) = "#" Then
                    sX = Mid(sX, 2)
                    If Not IsNumeric(sX) Or CLng(sX) < 0 Then
                        ErrMsg FATAL_ERR, "Bad Column in Adapter ='" & sX & "'"
                        End
                    End If
                ElseIf FF(PTRN_HDR, Col) = FOR_PROCESS Then
                  '-- �������� ��������� �������� - ���������� IsThisProc
                    If Proc = "" Then               '������ ������ ���
                        Proc = FF(PTRN_VALUE, Col)
                        sProc = Split(Proc, ",")
                        For iProc = LBound(sProc) To UBound(sProc)
                            If sX = PublicProcName Then
                                IsThisProc = True
                                ThisProcCol = Col + 1
                                If FF(PTRN_COLS, Col) = PublicProcName Then GoTo NextCol
                                ErrMsg FATAL_ERR, "������ � ������� '" _
                                    & DB_MATCH.Sheets(TOC).Cells(R.iTOC, TOC_FORMNAME) & "'" _
                                    & " � Col=" & Col & vbCrLf & "��������� ��� �������� '" _
                                    & PublicProcName & "', � �� '" & FF(PTRN_COLS, Col) & "'"
                                End
                            End If
                        Next iProc
                    End If
                    If Not IsThisProc Then GoTo NextRow
    '''            Else
    '''                iX = sX
    '''                X = Workbooks(R_To.RepFile).Sheets(R_To.SheetN).Cells(i, iX)
                End If
                
              '--- ����� ��������
                Rqst = FF.Cells(PTRN_ADAPT, Col)
                F_rqst = FF.Cells(PTRN_FETCH, Col)
    
                Y = Adapter(Rqst, X, F_rqst, IsErr, R_From.EOL, i, Col)
    
                If IsErr Then
                    iTo = iTo - 1
''                    Exit For
                    GoTo NextRow
                End If
              '--- ���������� � SheetNew �������� Y � ���������� ������� ������
                Width = Split(FF.Cells(PTRN_WIDTH, Col), "/")
                fmtCell Workbooks(R_To.RepFile), R_To.SheetN, Width, Y, iTo, Col
                
NextCol:            tot2(Col) = tot2(Col) + (Timer() - beg2(Col))   ' ��������������
            Next Col
        Next iPass
NextRow:
    Next i
'���� ������ � �������� NewSheet ��������� ������, ����� IsErr ��������=True - ������� ��� ������
    If IsErr And ToDoc <> "" Then Workbooks(R_To.RepFile).Sheets(R_To.SheetN).Rows(iTo).Delete

    ' ��������������
    tot1 = tot1 + (Timer() - beg1)
    profileStr = ""
    For Col = 1 To FF.Columns.Count
        profileStr = profileStr & " " & Format(tot2(Col), "###0.00")
    Next Col
    LogWr "adapt profile: total = " & Format(tot1, "###0.00") _
        & vbCrLf & "By steps = " & profileStr
End Sub
Function Adapter(ByVal Request As String, ByVal X As String, ByVal F_rqst As String, _
    ByRef IsErr As Boolean, Optional ByVal iPatternLine As Long, _
    Optional ByVal iRow As Long, Optional ByVal iCol As Long, _
    Optional ByVal PutToRow As Long) As String
'
' - Adater(Request, X, F_rqst, IsErr) - ������������ X � �������� "Request"
'    � �������� ������� � ��������� F_rqst. IsErr=True - ������ � ��������
' 4.9.12
' 6.9.12 - bug fix
'25.9.12 - Dec(CurRate)
' 3.10.12 - ������� GetCol � ����������� ' GetCol/1C.xlsx,�������,5/SF:2:11
'12.10.12 - ������� GoodType(X)
'14.10.12 - ������� OppFilter ��� ������� ���� Select
'18.10.12 - � OppFilter ��������� EOL
'23.10.12 - CopyToVal � CopyFrVal
'25.10.12 - ������� ����������, ���������� �� ������� ��������
'18.11.12 - ��������� ������ "�������"/"�������"
'19.11.12 - ���������� ��������� � ������� ������� ������ - � �.�. InvN
'29.12.12 - Compressor(X)
' 4.1.13 - ������� OppName ��� ��������; ��������� ���������� Array
' 5.1.13 - ������� <>"" � <>1; ��������� ��������� WP � ��������� ������
' 7.1.13 - ��������� � GoodType - ������ � ������� ��������, ��������, �����
'10.1.13 - ������� "�������; ����������� TypeSFopp
'23.1.13 - ����� �������� IsBalky � BalkyOppId
' 7.4.13 - �������� ��� ���: BTO_Date, BTO_Order, BTO_Ord
'23.4.13 - ������� GetInv1C
'11.8.13 - �������� "SN+" � "=Registered"
' 1.9.13 - ��� ��������� ByVal, ����� ������������� ByRef IsErr
' 9.9.13 - ��������� �������� TypOpp, LineOpp, KindOpp
'14.9.13 - �������� ������� AltFetch: Y=X ��� ��� ������������ �� Fetch, ���� ����
'16.9.13 - �������� Optional �������� PutToRow
'16.10.13 -bug fix - ������ ��������� � WP ��� �������� ContrK

    Dim FF() As String, Tmp() As String, InitX As String
    Dim i As Long, Par() As String, Z(10) As String
    Dim WP_Row As Long  ' ������ ��� ������ ��������� ���������, ����������� � Select
    Dim LocalTOC As TOCmatch
    
    IsErr = False
    X = Compressor(X)
    InitX = X
    
'--- ������ ������ �������� ���� <���>/C1,C2,C3...
    Dim AdapterName As String
    AdapterName = ""
    If Request <> "" Then
        Tmp = Split(Request, "/")
        AdapterName = Tmp(0)
        If InStr(Request, "/") <> 0 Then
            Par = Split(Tmp(1), ",")
'!!'            Call ArrayZ(Z, PAY_SHEET, iRow, Par)
        End If
    End If

'======== ������������� �������� ��� �������������� ��������� X ����� Fetch =========
    Select Case AdapterName
    Case "MainContract":
        X = Trim(Replace(X, "�������", ""))
    Case "<>0", "SN+":
        If X = "0" Then X = ""
    Case "<>"""""
        If X = "" Then
            IsErr = True
            Exit Function
        End If
    Case "=Registered"
        If X <> "Registered" Then
            IsErr = True
            Exit Function
        End If
    Case "ContrK":
        Dim MainDog As String
        MainDog = DB_1C.Sheets(PAY_SHEET).Cells(iRow, CLng(Par(0)))
        X = ContrCod(X, MainDog)
    Case "AltFetch":    ' Y=X ��� ��� ������������ �� Fetch, ���� ����
        Dim Str As String
        Str = FetchDoc(F_rqst, X, IsErr)
        If Str <> "" Then X = Str
        GoTo AdapterSelect
    End Select
    
'--- FETCH ������ ������ ���������� �� ���������� ���� <Doc1>/C1:C2,<Doc2>/C1:C2,...
    If F_rqst <> "" And X <> "" Then
        FF = Split(F_rqst, ",")
        For i = LBound(FF) To UBound(FF)
            X = FetchDoc(FF(i), X, IsErr)
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' ������ ������������ ������ ���� ��������� �� ����������� �� Doc ��������.
' � ���������� ���� ������������ ������ x(1 to 5) � ���������� � Fetch ��������� ���
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Next i
    End If

'******* ���������� �������� � ����������� Par ******
''    Set DB_TMP = FileOpen(F_TMP)
''    With DB_TMP.Sheets(WP)
AdapterSelect:

    Adapter = ""
'''''    If SheetWP = WP Then
'''''        Adapter = AdpterWP()
'''''        Exit Function
'''''    End If
    
    Select Case AdapterName
    Case "", "MainContract", "<>""""", "=Registered", "AltFetch": Adapter = X
    Case "��", "��������", "��������_�_SF", "Vendor":
        On Error GoTo AdapterFailure
        Adapter = WorksheetFunction.VLookup(X, DB_MATCH.Sheets("We").Range(AdapterName), Par(0), False)
        On Error GoTo 0
    Case "Dec": Adapter = Dec(X)
    Case "GetCol":
        If X <> "" Then           ' GetCol/1C.xlsx,�������,5 [/SF/2:11]
            Adapter = Workbooks(Par(0)).Sheets(Par(1)).Cells(CLng(X), CLng(Par(2)))
            If UBound(Tmp) > 1 Then
                Adapter = FetchDoc(Tmp(2) & "/" & Tmp(3), Adapter, IsErr)
            End If
        End If
    Case "GoodType": Adapter = GoodType(X)
    Case "GoodJob":
        Call ArrayZ(Z, PAY_SHEET, iRow, Par)
        If GoodJob(Z(1), X, CLng(Par(0))) Then Adapter = "1"
    Case "CurISO":  Adapter = CurISO(X)
    Case "CurRate": Adapter = Dec(CurRate(X))
    Case "����":    If X <> "" Then Adapter = DDMMYYYY(X)
    Case "ContrK":  Adapter = X '�������������� � ��� ContrCod � ��������������
    Case "SeekInv": Adapter = SeekInv(X)
    Case "InvN":
        Tmp = Split(X, " ")
        If UBound(Tmp) > 0 Then Adapter = Tmp(0)
    Case "SeekPayN":
        Dim Inv As String, Client As String
        Inv = ActiveSheet.Cells(iRow, CLng(Par(0)))
        Client = ActiveSheet.Cells(iRow, CLng(Par(1)))
        If Inv <> "" And IsDate(X) Then Adapter = SeekPayN(Inv, Client, X)
        If Adapter = "0" Then Adapter = ""
    Case "<>1":
        Call ArrayZ(Z, PAY_SHEET, iRow, Par)
        If Z(0) = "1" Then
            IsErr = True
        Else
            Adapter = X
        End If
    Case "SN+":
        Adapter = X & "+"
    Case "OppName":
        If X <> "" Then
            Adapter = X
        Else
            Call ArrayZ(Z, PAY_SHEET, iRow, Par)
            Adapter = Z(2) & "-" & Z(3) & " " & ContrCod(Z(4), Z(5))
        End If
    Case "BTO_Date":
        Adapter = Mid(X, 2, WorksheetFunction.FindB(" ", X) - 2)
    Case "BTO_Order":
        Dim ChBeg As Long, ChEnd As Long
        ChBeg = WorksheetFunction.FindB("�� �����", X) + 9
        ChEnd = WorksheetFunction.FindB(" ", X, ChBeg)
        Adapter = Mid(X, ChBeg, ChEnd - ChBeg)
    Case "TypeSFopp":
        Call ArrayZ(Z, PAY_SHEET, iRow, Par)
        If Z(0) = "1" Then Adapter = "��������"
        If Z(1) = "1" Then Adapter = "��������"
        If Z(3) = "1" Then Adapter = "������"
        If X = "����������" Then Adapter = X
        If X = "������������" Then Adapter = "������"
    Case "LineOpp":
        Adapter = "����������� ����������� (��)"
        X = GoodType(X)
        Select Case X
        Case "����������": Adapter = "��������� ��������� � ���"
        Case "������":     Adapter = "������"
        Case "������":     Adapter = "������"
        Case "������������": Adapter = "������"
        End Select
    Case "KindOpp":
        Dim GoodTp As String
        GoodTp = GoodType(X)
        On Error Resume Next
        Adapter = WorksheetFunction.VLookup(GoodTp, DB_MATCH.Sheets(We).Range("OppTypeRng"), 4, False)
        On Error GoTo 0
    Case "Max":
        Call ArrayZ(Z, PAY_SHEET, iRow, Par)
        Adapter = X
        For i = LBound(Z) To UBound(Z)
            Adapter = WorksheetFunction.Max(CLng(Adapter), CLng(Z(i)))
        Next i
    Case "EmptyBuddy":
        On Error GoTo AdapterFailure
        Adapter = WorksheetFunction.VLookup(X, DB_MATCH.Sheets("We").Range("��������"), 3, False)
        On Error GoTo 0
        If InStr(Adapter, X) = 0 Then Adapter = X
    Case "ForceTxt":
        Adapter = "'" & X
    Case "DogVal":                                      '=Max(�������, �����, ����� ��������)
        Dim Vpaid As Long, Vinv As Long, Vdog As Long   ' �������� �������, ����� � ��������
        Dim sDog As String, DogCur As String            ' ��� �������� � ��� ������
        LocalTOC = GetRep(PAY_SHEET)
        With Workbooks(LocalTOC.RepFile).Sheets(LocalTOC.SheetN)
            Vpaid = .Cells(iRow, PAYRUB_COL)
            Vinv = .Cells(iRow, PAYINVVAL_COL)
            sDog = .Cells(iRow, PAYDOGVAL_COL)
            DogCur = .Cells(iRow, PAYDOGCUR_COL)
        End With
        If sDog <> "" Then
            If Not IsNumeric(sDog) Then
                ErrMsg FATAL_ERR, "�� �������� �������� � ����� �������� �� " & X
                Stop
                End
            End If
            Vdog = CDbl(sDog) * CurRate(DogCur)
        End If
        Adapter = Dec(Application.Max(Vpaid, Vinv, Vdog, CLng(X)))
    Case "TypOpp":  '!!' ������� ��������
    ' -- ������������� ���� ������� �� ���� � ������������ ������
'        Dim Good As String
'        LocalTOC = GetRep(PAY_SHEET)
'        With Workbooks(LocalTOC.RepFile).Sheets(LocalTOC.SheetN)
'            Good = .Cells(iRow, PAYGOOD_COL)
'        End With
        Adapter = TypOpp(X)
    Case "CopyPattern", "OppType", "SetOppButton":
        Adapter = AdapterWP(AdapterName, X, Par, iRow, iCol, PutToRow, iPatternLine)
    Case "IsBalky":
        Call ArrayZ(Z, PAY_SHEET, iRow, Par)
        If Z(0) = "" Or Z(1) = "1" Or Z(2) <> "����������" Then
            IsErr = True
        Else
            Adapter = X
        End If
    Case "BalkyOppN":  'SFopp/4:2 � ��������� ������� OpportunityActivityKind = "����������"
        Dim BalkyExists As Boolean: BalkyExists = False
        Dim Rdoc As TOCmatch, N As Long, FromN As Long
        Rdoc = GetRep(SFopp)
        FromN = 1
        With Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN)
            Do
                N = CSmatchSht(X, SFOPP_ACC1C_COL, Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN), FromN)
                If N = 0 Then
                    If BalkyExists Then Exit Function
                    GoTo AdapterFailure
                End If
                Dim OppCloseDate As Date, DaysToClos As Long
                OppCloseDate = .Cells(N, SFOPP_CLOSEDATE_COL)
                DaysToClos = OppCloseDate - Now
                If .Cells(N, SFOPP_LINE_COL) = OPP_BALKY And DaysToClos > 365 Then
                    If BalkyExists Then
                    End If
                    BalkyExists = True
                    Adapter = .Cells(N, SFOPP_OPPN_COL)
                End If
                FromN = N + 1
            Loop
        End With
    Case Else
        If Left(AdapterName, 1) = """" Then
            Adapter = Mid(AdapterName, 2)
        Else
            ErrMsg FATAL_ERR, "Adapter> �� ���������� " & AdapterName
        End If
    End Select
''    End With
    Exit Function
    
AdapterFailure:
    ErrMsg WARNING, "������� " & AdapterName & "(" & X & ") �� ������� ������"
SkipLine:
    IsErr = True
End Function
Function AdapterWP(AdapterName, X, ByRef Par, _
    ByVal iRow As Long, ByVal iCol As Long, _
    ByVal PutToRow As Long, ByVal iPatternLine As Long) As String
'
' - AdapterWP() - ��������� ��������� ��� �������� WP
' 5.1.2013
' 21.10.13 - ����������

    Dim i As Long, Z(10) As String, WP_Row As Long
     
    Set DB_TMP = FileOpen(F_TMP)
    With DB_TMP.Sheets(WP)
    
        Select Case AdapterName
        Case "CopyPattern":
            .Cells(iPatternLine, iCol).Copy .Cells(PutToRow, iCol)
        Case "OppType":             ' ������������� ���� �������
        ''''        Call ArrayZ(Z, PAY_SHEET, iRow, Par)
            If X = "������������" Then X = "������"
            If X = "Autodesk" Then
        ''  !          If IsSubscription(.Cells(WP_PAYMENT_LINE, CLng(Par(0))), X) Then
                If IsSubscription(Z(0), X) Then
                    X = "��������"
                Else
                    X = "��������"
                End If              '!' �� ��������������� ������!!
            End If
            AdapterWP = X
        ' ������� ���� ������������ ������, ����� ������ � ���������, � �� ������ � ��������
                Dim Rdoc As TOCmatch, Doc As String
''        !Doc = .Cells(iRow, 1)
                Rdoc = GetRep(Doc)
'!!!!!                AdapterWP = CSmatchSht(X, SFOPP_OPPID_COL, Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN))
'''        !            .Cells(iRow + PTRN_LNS, 11) = "�������"
'''        !            .Cells(iRow + PTRN_LNS, 11).Interior.Color = rgbBlue
'''        !            If Adapter = .Cells(iRow + 1, 4) Then Adapter = "-1"
'            End If
        Case "SetOppButton":
            Const PAY_DOG_COL = 5
            '---------------------------------------------
            ' -- ������� ������� --         -������-
            ' ����������� "�������" (="")   "�������"
            ' ���� �� �� ������ � ��������  "�������"
            ' ���� � �� ������ � ��������   "�������"
            '---------------------------------------------
            ' ������ ����� �� ������������� ����� � ������ �������
            .Cells(26, iCol).Copy .Cells(PutToRow, iCol)
            If .Cells(8, 7) = "" Or .Cells(8, 10) <> "" Then
                .Cells(PutToRow, iCol).Value = "������� ->"
            End If
        End Select
    End With
End Function
Sub testArrayZ()
    Dim Z(3) As String, Par(3) As String
    Par(0) = "9"
    Par(1) = "31"
    Par(2) = "25"
    Par(3) = "26"
    Dim t0, t1, t2
    t0 = Now
    Call ArrayZ(Z, PAY_SHEET, 2886, Par)
    t1 = Now
    Call ArrayZ(Z, PAY_SHEET, 3082, Par)
    t2 = Now
    t2 = 1000 * (t2 - t1)
    t1 = 1000 * (t1 - t0)
End Sub
Sub ArrayZ(ByRef Z() As String, Doc As String, iRow As Long, ByRef Par() As String)
'
' - ArrayZ(Z(), Doc, iRow, Par())   - �� ������� ������� ������� � Par ���������� � ������ Z()
'                               �������� �� ����� ��������� Doc � ������ iRow � �������� Par()
' 4.1.13

    Dim R As TOCmatch, i As Long
    
    R = GetRep(Doc)
    With Workbooks(R.RepFile).Sheets(R.SheetN)
        For i = LBound(Par) To UBound(Par)
            Z(i) = Trim(.Cells(iRow, CLng(Par(i))))
        Next i
    End With
End Sub
Sub TestX_Parse()
'
' T TestX_Parse()
'
'   8.9.13

    Dim X

    X = X_Parse(X_rqst:="7", InDoc:=PAY_SHEET) ' ������ ������� Date �� EOL ��������
    X = X_Parse(X_rqst:="#1", OutDoc:=NEW_OPP) ' ������ ������� �������� �� ������� 1 EOL NewOpp
' ������ ������� �������� �� ������� 4 ��������� ��������� ����� NewOpp
    X = X_Parse(X_rqst:="#4", OutDoc:=NEW_OPP, iLine:=1)
    X_Parse X_rqst:=EXT_PAR ' ������ ���� FATAL_ERR, � �� ��� ���������� ""
    X_Parse X_rqst:=EXT_PAR, ExtPar:="����������� �����"    ' ���������� ������ �� ExtPar
End Sub
Function X_Parse(Optional ByVal InDoc As String, _
    Optional ByVal OutDoc As String = "", _
    Optional ByVal X_rqst As String, _
    Optional ByVal iLine As Long, _
    Optional ByVal PutToRow As Long, _
    Optional ByVal ExtPar As String, _
    Optional ByVal ExtPar2 As String, _
    Optional ByVal ExtPar3 As String) As String
'
' -  X_Parse([InDoc,OutDoc,X_rqst,iLine,PutToRow,ExtPar)  - ������ ������ �-��������� ��������
'            �� ���� X_rqst �������. �� �������� # � X_rqst X ����������� �� OutDoc
'   �����
'           X_rqst          - ����������� ������ ��������� �������� X
'         (iLine,iCol)      - ����� ������ ������� WP ��� �������, ������ �� ����� �������.
'                             ���� ������� (iLine, iCol), �� �������� X_rqst ����������� �� WP
'       [PutToRow,PutToCol] - ����� ������, ���� ��������� ��������� �������� � WP
'           [iLine]         - ����� ������ ��� ������� WP ���� iLine
'
' � ���� ������� �������� ����������� #6 ��� !6, � ����� "ExtPar"
'  * ���� # ��������, ��� ���������� �� ������� � ActiveSheet, � ������� ������ �������
'  * ���� & - ��� � ���� ������� �� ������� -- �� ��������
'  * ���� ! - ��� � ���� ������� �� ������ WProw ���� �������
'  * "ExtPar" - ���������� X_Parse �������� ��������� ExtPar
'  ? ����� �������� ����������� ���� #3/Format, ��� ����� "/" ������� ���.����������
'
' 22.10.12
' 25.10.12 - �c��������� � ����� � HashFlag=True
' 11.11.12 - �������� ��������� !<Col> ��� ��������� WProw
' 15.11.12 - Optional ���������
'  8.09.13 - ��������� ������ ������� �� Optional �������� X_rqst, �� ������ �� WP
'          - Optional �������� ExtPar: ���� � ���� ������� "ExtPar", �=����� ���������
' 24.09.13 - ��������� ExtPar2
' 20.10.13 - ��������� ����������� "&" � Columns

    Dim InDocTOC As TOCmatch, OutDocTOC As TOCmatch
    Dim ErrStr As String, FF As Range
    Dim sX() As String
    Dim iX As Long, WP_Row As Long
    Dim RefType As String
    Dim iCol As Long            '= ����� �������
    
    X_Parse = ""
    If IsMissing(InDoc) Then InDoc = ActiveSheet.Name
    If IsMissing(X_rqst) Or X_rqst = "" Then GoTo Ex
    ErrStr = "X_Parse> "
    
    If IsNumeric(X_rqst) Then
        If Trim(InDoc) = "" Then GoTo ErrInDoc
        InDocTOC = GetRep(InDoc)
        If IsMissing(iLine) Then iLine = InDocTOC.EOL
        iCol = X_rqst
        GoTo GetX
    ElseIf X_rqst = EXT_PAR Then
        If IsMissing(ExtPar) Then GoTo ErrExtPar
        X_Parse = ExtPar
        GoTo Ex
    ElseIf X_rqst = EXT_PAR_2 Then
        If IsMissing(ExtPar2) Then GoTo ErrExtPar
        X_Parse = ExtPar2
        GoTo Ex
    ElseIf X_rqst = EXT_PAR_3 Then
        If IsMissing(ExtPar3) Then GoTo ErrExtPar
        X_Parse = ExtPar3
        GoTo Ex
    Else
        RefType = Left(X_rqst, 1)
        If RefType = "#" Or RefType = "&" Then
            iCol = Mid(X_rqst, 2)
            If IsMissing(OutDoc) Then OutDoc = InDoc
'''' ---- �� �������� ----------
''''            If RefType = "&" Then
''''                With InDocTOC
''''                    i = InDocTOC.iTOC
''''                    X_Parse = Workbooks(.RepFile).Sheets(.SheetN).Cells(, iCol)
''''                End With
''''                GoTo Ex
''''            End If
            If IsMissing(PutToRow) Or Not IsNumeric(PutToRow) Then GoTo ErrPutToRow
            OutDocTOC = GetRep(OutDoc)
            With OutDocTOC
                X_Parse = Workbooks(.RepFile).Sheets(.SheetN).Cells(PutToRow, iCol)
            End With
        End If
        GoTo Ex
    End If
    
GetX:
    If Not IsNumeric(iLine) Or iLine <= 0 Or iLine > InDocTOC.EOL Then GoTo Err_iLine
    If Not IsNumeric(iCol) Or iCol <= 0 Then GoTo Err_iCol
    
    X_Parse = Workbooks(InDocTOC.RepFile).Sheets(InDocTOC.SheetN).Cells(iLine, iCol)

Ex: Exit Function
ErrInDoc:       ErrMsg FATAL_ERR, ErrStr & "���������� InDoc"
ErrOutDoc:      ErrMsg FATAL_ERR, ErrStr & "���������� OutDoc"
ErrExtPar:      ErrMsg FATAL_ERR, ErrStr & "���������� ExtPar"
Err_iLine:      ErrMsg FATAL_ERR, ErrStr & "�������� iLine=" & iLine
Err_iCol:       ErrMsg FATAL_ERR, ErrStr & "�������� iCol=" & iCol
ErrPutToRow:    ErrMsg FATAL_ERR, ErrStr & "�������� PutToRow=" & PutToRow
End Function
Function X_ParseWP(ByVal iRow, ByVal iCol, ByVal PutToRow, putToCol, _
    Optional iLine, Optional ExtPar As String) As String
'
' -  X_ParseWP([InDoc,OutDoc,X_rqst,iLine,PutToRow,ExtPar)  - ������ ������ � ��� WP
'
'''''                X = X_Parse(iRow, iCol, putToRow, putToCol, iLine)

    Dim WP_TOC As TOCmatch, InDoc As String
    Dim iX As Long, WP_Row As Long
    Dim PtrnType As String
    Dim X_rqst As String, sX() As String
    
    WP_TOC = GetRep(WP)
    With Workbooks(WP_TOC.RepFile).Sheets(WP_TOC.SheetN)
        InDoc = .Cells(iRow, 1)
        WP_Row = iRow - 1 + PTRN_VALUE
        PtrnType = .Cells(iRow, 2)
        If PtrnType = PTRN_SELECT Then WP_Row = iRow + PTRN_LNS + .Cells(iRow + 3, 3) - 1
        PutToRow = WP_Row: putToCol = iCol
        X_rqst = .Cells(iRow - 1 + PTRN_COLS, iCol)
    End With
    
GetX:
    X_ParseWP = X_Parse(InDoc:=InDoc, OutDoc:=WP, X_rqst:=X_rqst, iLine:=iLine, PutToRow:=WP_Row)

'    If X_rqst = "" Then GoTo Ex
'
'    sX = Split(X_rqst, "/")
'    RefType = Left(sX(0), 1)
''''    If OutputDoc = WP Then      '���� OutputDoc �� ������, ��� ���������� ������ � WP
''''    If RefType = "#" Or RefType = "!" Then sX(0) = Mid(sX(0), 2)
'    If RefType = "#" Or RefType = "!" Then sX(0) = Mid(sX(0), 2)
'    X_Parse = sX(0)
''        If IsMissing(Form) Then GoTo ErrInDocForm
''        Set FF = DB_MATCH.Sheets(Header).Range(Form)
''        InDoc = FF(DB_MATCH.Cells(1, 1))
''''
''''    iX = 0
''''    If UBound(sX) >= 0 Then iX = sX(0)
''''    If iX > 0 Then
''''        Select Case PtrnType
''''        Case "������", "������": GoTo GetFromWP
''''        Case "iLine":
''''            WP_Row = iLine
''''            GoTo GetFromActiveSheet
''''        Case PTRN_SELECT:
''''            WP_Row = .Cells(putToRow, 5)
''''            GoTo GetFromActiveSheet
''''         Case Else:
''''            ErrMsg FATAL_ERR, "WP_Adapt> �������� ��� ������� " & PtrnType
''''        End Select
''''        End If
''''
''''GetFromWP:
''''        If iX > 0 Then X_Parse = .Cells(WP_Row, iX)
''''        GoTo Ex
''''    End With
''''
''''GetFromActiveSheet:
''''    If RefType = "!" Then
''''        WP_Row = putToRow
''''        GoTo GetFromWP
''''    ElseIf RefType = "#" Then
''''        WP_Row = iRow + PTRN_VALUE - 1
''''        GoTo GetFromWP
''''    End If
''''    If iX > 0 Then X_Parse = ActiveSheet.Cells(WP_Row, iX)
'    If OutputDoc = WP Then      '���� OutputDoc �� ������, ��� ���������� ������ � WP
'        GetRep WP
'        With DB_TMP.Sheets(WP)
'            WP_Row = iRow - 1 + PTRN_VALUE
'
'            PtrnType = .Cells(iRow, 2)
'            If PtrnType = PTRN_SELECT Then WP_Row = iRow + PTRN_LNS + .Cells(iRow + 3, 3) - 1
'
'            putToRow = WP_Row: putToCol = iCol
'
'            X_rqst = .Cells(iRow - 1 + PTRN_COLS, iCol)
'        End With
'    End If
'    X_ParseWP = XParse(OutputDoc:=WP, iRow:=iRow, iCol:=iCol)
    End Function
Function FetchDoc(F_rqst, X, IsErr, Optional ByRef FromN As Long = 1) As String
'
' - FetchDoc(F_rqst, X, IsErr) - ���������� ������ �� ���������� ���������
'                   �� ������� F_rqst ��� �������� ���� X. IsErr=True - ������
'
' * F_rqst ����� ��� <Doc>/C1[:C2][/W]
' * <Doc>   - ��� ���������, ����� ����������� ������
' *   /     - �������� ������ ����������. �������� ��������� �����.
' *   :     - ��������� ��������� ������ ������
' *             ������ ������ - ���������� ��������� ��� ���������� �1[:�2]
' *  C1                 ���� ���� ������ �1 - ����������� ���� ����� �1
' * C1:C2               ���� �1:�2 - Lookup �� �1 -> �� C2 � Range �� Doc
'                               C2 ����� ����� �������� �, ����� �� �1 ������������
'                                  ����� ������, ��� ����� �������� �1
' *             ������ ������ - ��������� ��������� ������ Fetch - /W ��� /0
' *  /W             - WARNING � Log, ��������� IsErr=False, ���� ��������� ""
' *  /0             - "" ������ ��������� (��������, ������� � ������)
' *  /D             - Default - "" ��������, �� IsErr=True ��� �������� �� ���������
' *             ������ ������ ����������� - �������� Log � IsErr = True
'
' 5.9.12
' 14.9.12 - �������� /D ��� ������ ������ - "�� ���������"
' 4.11.12 - Fetch ���������� ����� ������ � ������ <Doc>/C1:�
' 3.09.13 - Optional FromN ��������� ����� ����� �� � ������ ���������
' 20.10.13 - bug fix - /0 � X="" �� ������ ������ IsErr=True

    FetchDoc = ""
    If F_rqst = "" Or X = "" Then GoTo ErrExit
        
    Dim Tmp() As String, Cols() As String, S As String
    Dim Doc As String, C1 As Long, C2 As Long, Rng As Range, N As Long
            
    Tmp = Split(F_rqst, "/")
    Doc = Tmp(0)
    Cols = Split(Tmp(1), ":")
    C1 = Cols(0)
    
    Dim Rdoc As TOCmatch, W As Workbook
    Rdoc = GetRep(Doc)
    
    If UBound(Cols) < 1 Then
'--- �������� �1 - � ������ ���� �������� - ��������� �������� �� �������
        Dim Indx As Long
        Indx = X
'!!!!!!!!!!!!!!!!!!!!!!!!!!!
' ������ Indx=� - ��� ������ �����, �� � ���������� ��� ���� split
'!!!!!!!!!!!!!!!!!!!!!!!!!!!
        If Indx <= 0 Then
            ErrMsg WARNING, "FetchDoc: " & Doc & "(" & Indx & "," & C1 _
                & ") - ������������ ����� ������"
            GoTo ErrExit
        End If
        S = Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN).Cells(Indx, C1)
    Else
'--- �������� �1:C2 - � ������ 2 ��������� - ��������� �������� �� Lookup ��� �
        If IsNumeric(Cols(1)) Then C2 = Cols(1)
        S = ""
        N = CSmatchSht(X, C1, Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN), FromN)
        FromN = N
        If N <> 0 Then
            If Cols(1) = "�" Then
                S = N
            Else
                S = Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN).Cells(N, C2)
            End If
        End If
    End If
'--- ��������� ������ 2 -- ���� S=""
    If S = "" Then
        If UBound(Tmp) >= 2 Then
            If Tmp(2) = "W" Then
                ErrMsg WARNING, "�������> ������ " & F_rqst _
                    & "(" & X & ") �� ��������, ��������� <�����>"
            End If
            If Tmp(2) <> "0" Then GoTo ErrExit
        Else
            ErrMsg WARNING, "�������> ������ " & F_rqst _
               & "(" & X & ") �� ��������, ��������� <�����>"
            GoTo ErrExit
        End If
    Else
        FetchDoc = S
    End If
    
OK_Exit:    IsErr = False
    Exit Function
ErrExit:    If Right(F_rqst, 2) = "/0" Then GoTo OK_Exit
            IsErr = True
End Function
Sub testfmtCell()
'   ����� fmtCell()
' 17.12.12 - �������� ���� ������ �������
    
    Dim fmt(0 To 1) As String
    Set DB_TMP = FileOpen(F_TMP)
    
    fmt(1) = "Dbl"
    fmtCell DB_TMP, "NewOpp", fmt, "3m3", 2, 2
    fmtCell DB_TMP, "NewOpp", fmt, 33333.3, 2, 2
    
    fmt(1) = "Txt"
    fmtCell DB_TMP, "NewOpp", fmt, "xxx", 2, 2
    fmt(1) = "@"
    fmtCell DB_TMP, "NewOpp", fmt, "yyy", 2, 2
    
    fmt(1) = "Date"
    fmtCell DB_TMP, "NewOpp", fmt, "1/2/2012", 2, 2
    fmt(1) = "# ##0.0000"
    fmtCell DB_TMP, "NewOpp", fmt, 5666, 2, 2
    fmt(1) = "# ##0.00"             ' �������� ��� ������ ~ Dbl
    fmtCell DB_TMP, "NewOpp", fmt, 5666, 2, 2
    fmt(1) = "0.00"
    fmtCell DB_TMP, "NewOpp", fmt, 5666, 2, 2
    fmt(1) = "0"
    fmtCell DB_TMP, "NewOpp", fmt, 22, 2, 2
    fmt(1) = "0%"
    fmtCell DB_TMP, "NewOpp", fmt, 5.666, 2, 2
    fmt(1) = "0.00%"
    fmtCell DB_TMP, "NewOpp", fmt, 5.666, 2, 2
    Stop
End Sub
Sub fmtCell(ByVal db As Workbook, ByVal list As String, fmt() As String, _
            ByVal Value, ByVal PutToRow As Long, ByVal putToCol As Long)
'
'   - fmtCell - ��������� ������� � ������ width ���� 0/Txt ��� 10/@
'
' db � list - ����� Workbook � �����, ���� �������� ..
' value     - �������� � ���� ..
' fmt(1)    - ������� ������ [putToRow,putToCol]
'
'   * ������������ ������� Txt, Dbl, Date � ����� ������, ���������� Excel
'
'  7.12.12 - ������� ������� ������ "Dbl", "Txt", "Date" � ������ "width" � sub WP_Adapt
'  8.12.12 - ������ ��o��������� ������ � ������ width
' 17.12.12 - �������� ���� ������ ������� � testfmtCell()
' 19.12.12 - ������� ����������� ����� � Dbl � testfmtCell()
' 17.12.12 - �������� ���� ������ �������
' 12.9.13 - ��������� ���������� ����� ��� Dbl
' 16.9.13 - ��������� �������� 0% � Dbl
' 20.10.13 - ����������� ��������� Txt ��� ������, ������� �� ����

    If UBound(fmt) > 0 Then
        If fmt(1) = "Dbl" Then
            Value = CDbl(Value)
            db.Sheets(list).Cells(PutToRow, putToCol).NumberFormat = "# ### ##0.00"
        ElseIf fmt(1) = "Date" Then
            db.Sheets(list).Cells(PutToRow, putToCol).NumberFormat = "[$-409]d-mmm-yyyy;@"
        ElseIf fmt(1) = "Txt" Then
            Value = "'" & Value
            db.Sheets(list).Cells(PutToRow, putToCol).NumberFormat = "@"
       ElseIf fmt(1) = "0%" Then
            Value = CDbl(Value) / 100
            db.Sheets(list).Cells(PutToRow, putToCol).NumberFormat = "0%"
        Else
            db.Sheets(list).Cells(PutToRow, putToCol).NumberFormat = fmt(1)
        End If
    End If
    db.Sheets(list).Cells(PutToRow, putToCol) = Value
End Sub
Sub testAdaptPass0()
'
' T testAdaptPass0()
'
' 23.9.13

    Dim Res() As Long
    Res = AdaptPass0("��� ������")
    Res = AdaptPass0("HDR_1C_Payment_MyCol")
End Sub
Function AdaptPass0(ByVal FormName As String) As Long()
'
' - AdaptPass0(FormName)    - ���������� ������ ������� ������� ������� FormName
'                             ��� ����������� � �������������� ������� Pass0
'
'   * � ������� 0 ��������������� ������� ����� ������� - ������ ������� �������
'   * � ��������� �������� - ������ ������� ��� ��������������� ������� Adapt
'
' 23.09.13

    Dim i As Long, FF As Range
    Dim IsAddForm As Boolean: IsAddForm = False
    Dim Arr() As Long, ArrLng As Long
    
    ReDim Arr(0) As Long:   Arr(0) = 0
    On Error GoTo Ex
        Set FF = DB_MATCH.Sheets(Header).Range(FormName & "_Pass0")
    On Error GoTo 0
    With FF
        ArrLng = .Columns.Count - 1
        ReDim Arr(ArrLng) As Long
        Arr(0) = ArrLng
        For i = 1 To ArrLng
            Arr(i) = CLng(.Cells(1, i + 1))
        Next i
    End With
Ex: AdaptPass0 = Arr
End Function
