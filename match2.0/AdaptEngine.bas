Attribute VB_Name = "AdaptEngine"
'---------------------------------------------------------------------------------------
' AdaptEngine - ��������� ���������.
'       * ������� - �������������� �������, ��������������� ��� ��������� ���� ���������
'       * �������� ��������� �������� � �������, � ��� �����:
'           1.(���.1) �������� �������������� �������� - �� ���������� ActiveSheet
'           2.(���.2) ��� ������� (������������ � �������� WP)
'               2.1 ������ (New, ->, Stop)
'               2.2 iLine - ������ �� ������ ����� iLine, ���������� ��� �������� xAdapt
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
'
' 19.11.12 �.�.�������, �.����
'   ������� ������:
' 11.11.12 - ��������� AdaptEngine �� ProcessEngine
'
' - WrNewSheet(SheetNew, SheetDB, DB_Line[,IdOpp]) - ���������� ����� ������
'                               � ���� SheetNew �� ������ DB_Line ����� SheetDB
'...........................
' - xAdapt(F, iLine) - ��������� �������� �� ����� F, ����������� ������ � ������
'                      �� ������ ����� iLine � ActiveSheet. ��������������� �� End
' * xAdapt_Continue(Button) - ����������� ������ xAdapt ����� ������� ������ Button
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

Const PTRN_VALUE = 2 ' �������� ������ - �������� - Value � �������
Const PTRN_WIDTH = 3 ' �������� ������ - ������ ������� � �������
Const PTRN_COLS = 4  ' �������� ������ ������� ������� � �������
Const PTRN_ADAPT = 5 ' �������� ������ ������ ��������� � �������
Const PTRN_FETCH = 6 ' �������� ������ ������ Fetch - ���������� �� ���-� � �������
Const PTRN_LNS = 6   ' ���-�� ����� � ������� �� ������ ������ ����� �� ������

Const PTRN_SELECT = "Select"

Sub WrNewSheet(SheetNew As String, SheetDB As String, DB_Line As Long, _
    Optional ExtPar As String)
'
' - WrNewSheet(SheetNew, SheetDB, DB_Line[,IdOpp]) - ���������� ����� ������
'                               � ���� SheetNew �� ������ DB_Line ����� SheetDB
'   * ��� � ��������� ��� ��������� ���������� � ������� � ���� ��������� �����.
'     ��� ������ �������� � Range � ������ "HDR_" & SheetNew � Forms ��� Headers
'   * ��������� � �������� ����� ��� <�����������>/<���1>,<���2>...
'   * � ������ ����� ��� ��������� ����� ������� ��������� �� ������� ����������
'   * ���� � ������� � ������ PTRN_COLS ������� "ExtPar", ���������� �������
'                                              �������� ExtPar = IdOpp
' 6.9.2012
' 26.10.12 - ��������� "�������" ������ � DB_TMP
' 27.10.12 - ������������� TOCmatch ��� "�������" ������
' 28.10.12 - �������� SheetDB - ���������� � ���� String

    Dim Rnew As TOCmatch, Rdoc As TOCmatch
    Dim P As Range
    Dim i As Long
    Dim X As String         '= �������������� �������� � SheetDB
    Dim sX As String        '���� � ������ PTRN_COLS �������
    Dim Y As String         '= ��������� ������ ��������
    Dim IsErr As Boolean    '=True ���� ������� ��������� ������
    
    Rnew = GetRep(SheetNew)
    Rnew.EOL = EOL(Rnew.SheetN, DB_TMP) + 1
    Rnew.Made = "WrNewSheet"
    Rdoc = GetRep(SheetDB)
    
    
    
    With DB_TMP.Sheets(SheetNew)
        Set P = DB_MATCH.Sheets(Header).Range("HDR_" & SheetNew)
        For i = 1 To P.Columns.Count
            sX = P.Cells(PTRN_COLS, i)
            If sX <> "" Then
                If sX = EXT_PAR Then
                    X = ExtPar
                Else
                    X = Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN).Cells(DB_Line, CLng(sX))
                End If
                
                Y = Adapter(P.Cells(PTRN_ADAPT, i), X, P.Cells(PTRN_FETCH, i), IsErr)
                
                If IsErr Then
                    .Rows(Rnew.EOL).Delete
                    Exit For
                Else
                    .Cells(Rnew.EOL, i) = Y
                End If
            Else
                .Cells(Rnew.EOL, i) = P.Cells(2, i) '!!!!!!!!!!!!!???????????!!!!!!!!!!!!
            End If
        Next i
    End With
    If Not IsErr Then
        RepTOC = Rnew
        WrTOC
    End If
End Sub
Sub xAdapt(F As String, iLine As Long)
'
' - xAdapt(F, iLine) - ��������� �������� �� ����� F, ����������� ������ � ������
'                      �� ������ ����� iLine � ActiveSheet
'   21.10.12
'   23.10.12 - X_Parse ������� � ��������� ������������
'    2.11.12 - ����� NewOpp ���� Select �� ����� �� ������ �������
'    9.11.12 - ������ � Named Range WP
'   11.11.12 - ������ ���������� ���� ��� ������� TraceWidth

    Const WP_PROTOTYPE = "WP_Prototype"

    Dim R As TOCmatch                           ' �������������� ��������
    Dim iRow As Integer, iCol As Integer        ' ������ � ������� ������� F
    Dim PtrnType As String                      ' ���� ��� �������
    Dim PutToRow As Long, PutToCol As Long
    Dim X As String                             ' �������� ��������
    Dim Rqst As String                          ' ������ - ��������� � ��������
    Dim F_rqst As String                        '
    Dim Y As String
    Dim IsErr As Boolean
    Dim iSelect As Long     '''', WP_Row As Long
    Dim i As Long
    Dim WP_Prototype_Lines As Long
            
'---- ������� ������ ���� WP
    Set DB_TMP = FileOpen(F_TMP)
    With DB_TMP
        Application.DisplayAlerts = False
        On Error Resume Next
        .Sheets(WP).Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
        DB_MATCH.Sheets(WP_PROTOTYPE).Copy Before:=.Sheets(1)
        .Sheets(1).Name = WP
    End With
'===== ��������� WP
    With DB_TMP.Sheets(WP)
        .Tab.Color = rgbBlue
        For i = 1 To EOL(WP, DB_TMP)
            .Rows(1).Delete
        Next i
        
        Dim FF As Range:  Set FF = DB_MATCH.Sheets(WP_PROTOTYPE).Range(F)
        FF.Copy .Cells(1, 1)
        .Cells(1, 5) = "'" & DirDBs & F_MATCH & "'!xAdapt_Continue"
'---- ������ ������ � ��������� ����������� �������
        For i = 1 To FF.Columns.Count
            If Not TraceWidth Then .Columns(i).ColumnWidth = FF.Cells(3, i)
        Next i
        
        .Cells(WP_CONTEXT_LINE, WP_CONTEXT_COL) = iLine
        WP_Prototype_Lines = EOL(WP, DB_TMP)
        For iRow = 1 To WP_Prototype_Lines Step PTRN_LNS
            PtrnType = .Cells(iRow, 2)
            
            R.EOL = -1                      ' ������������� EOL �� ������ Select
            If .Cells(iRow, 1) <> "" Then
                R = GetRep(.Cells(iRow, 1))
                Workbooks(R.RepFile).Sheets(R.SheetN).Activate
            End If
        '--- SelectLoop
            iSelect = 2
            Do
                For iCol = 5 To .UsedRange.Columns.Count
                    X = X_Parse(iRow, iCol, PutToRow, PutToCol, iLine)
                    
                    Rqst = .Cells(iRow - 1 + PTRN_ADAPT, iCol)
                    F_rqst = .Cells(iRow - 1 + PTRN_FETCH, iCol)
                    
                    Y = Adapter(Rqst, X, F_rqst, IsErr, R.EOL, iRow, iCol)
                    
                    If InStr(Rqst, "OppFilter") <> 0 And Y = "-1" Then GoTo OppEOL
                    X = .Cells(iRow + PTRN_COLS - 1, iCol)
                    If X = "-1" Then Exit For
                    fmt = .Cells(iRow + PTRN_WIDTH - 1, iCol)
                    If Not IsErr And X <> "" Then
                        .Cells(PutToRow, PutToCol) = Y
                        If fmt = "Dbl" And IsNumeric(Y) Then
                            Dim YY As Double
                            YY = Y
                            .Cells(PutToRow, PutToCol) = YY
                            .Cells(PutToRow, PutToCol).NumberFormat = "#,##0.00"
                        ElseIf fmt = "Date" Then
                            .Cells(PutToRow, PutToCol).NumberFormat = "[$-409]d-mmm-yyyy;@"
                        ElseIf fmt = "Txt" Then
                            .Cells(PutToRow, PutToCol).NumberFormat = "@"
                        End If
                    End If
                Next iCol
                If PtrnType = PTRN_SELECT Then
                    iSelect = .Cells(iRow + CLng(.Cells(iRow + 3, 3)) + 5, 5)
                    If iSelect < 0 Then Exit Do         ' ����� �� EOL ������������ ���������
                    .Cells(iRow - 1 + PTRN_VALUE, 4) = iSelect
                    .Cells(iRow - 1 + PTRN_COLS, 3) = .Cells(iRow - 1 + PTRN_COLS, 3) + 1
                    .Rows(iRow - 1 + PTRN_VALUE).Hidden = True
                End If
'''                WP_Row = WP_Row + 1
                                                ' ��� ������� Select ����� �� ����� ����������
            Loop While PtrnType = PTRN_SELECT   '.. �� ���������� ��������� OppFilter EOL SFopp
                
OppEOL:     .Rows(iRow - 1 + PTRN_COLS).Hidden = True
            .Rows(iRow - 1 + PTRN_ADAPT).Hidden = True
            .Rows(iRow - 1 + PTRN_WIDTH).Hidden = True
            .Rows(iRow - 1 + PTRN_FETCH).Hidden = True
        Next iRow
    End With
    DB_TMP.Sheets(WP).Activate
    
    If iSelect = 2 And Y = "-1" Then
       xAdapt_Continue "NewOpp", 1
    End If

'''''''''''''''''''''''''''''''''''
    End '''  ��������� VBA ''''''''
'''''''''''''''''''''''''''''''''''
End Sub
Sub xAdapt_Continue(Button As String, iRow As Long)
'
' * xAdapt_Continue(Button) - ����������� ������ xAdapt ����� ������� ������ Button
'                             ���� ���������� ���������� �� WP_Select_Button.
' 8/10/12
' 20.10.12 - ��������� ������ "�������"
' 10.11.12 - bug fix - ����������� ����� WP � ������������ Namer Range

    Dim Proc As String, Step As String, iStep As Long
    Dim iPayment As Long, OppId As String
        
'---- ��������� ��������� �� ����� WP, �� ���� ������ �������, ������� -----
    With ActiveSheet
        iPayment = .Cells(WP_CONTEXT_LINE, WP_CONTEXT_COL)
        OppId = .Cells(iRow, 6)
    End With
    
    If DB_TMP Is Nothing Then Set DB_TMP = FileOpen(F_TMP)
    If DB_1C Is Nothing Then Set DB_1C = FileOpen(F_1C)
    If DB_MATCH Is Nothing Then Set DB_MATCH = FileOpen(F_MATCH)
    
    With DB_MATCH.Sheets(Process)
        .Activate
        Proc = .Cells(1, PROCESS_NAME_COL)
        Step = .Cells(1, STEP_NAME_COL)
        iStep = ToStep(Proc, Step)
        .Cells(iStep, PROC_PAR2_COL) = iPayment + 1
    End With
    
    Select Case Button
    Case "STOP":
        StepOut Step, iStep
        ProcStart Proc
        End
    Case "->":
    Case "NewOpp":
        WrNewSheet NEW_OPP, WP, WP_PAYMENT_LINE
    Case "NewAcc":
    ' ���� �� ��������
'-------- ��������� ������ �� ������� ����� Select
    Case "�������":
        WrNewSheet NEW_PAYMENT, PAY_SHEET, iPayment, OppId
    Case "�������  ->"
        MS "->"
        Stop
        WrNewSheet DOG_UPDATE, PAY_SHEET, iPayment
    End Select
    
NextWP:         ProcStart Proc

End Sub
Sub Adapt(F As String)
'
' S Adapt(F) - ��������� �������� �� ����� F, ����������� ������ �� ���������
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

    StepIn
    
    Dim FF As Range     '= ����� F
    Dim R As TOCmatch
    Dim Rqst As String, F_rqst As String, IsErr As Boolean
    Dim X As String, Y As String
    Dim i As Long, Col As Long, iX As Long
''    Dim PutToRow As Long, PutToCol As Long
    
    Set FF = DB_MATCH.Sheets(Header).Range(F)
    
    With ActiveSheet
        R = GetRep(.Name)
        For i = 2 To R.EOL
            Progress i / R.EOL
            For Col = 1 To FF.Columns.Count
'''                X = X_Parse(i, Col)
'''                Rqst = FF.Cells(PTRN_ADAPT, Col)
'''                F_rqst = FF.Cells(PTRN_FETCH, Col)
'''
'''                Y = Adapter(Rqst, X, F_rqst, IsErr)
'''
'''                If Not IsErr Then .Cells(i, Col) = Y
                iX = FF(PTRN_COLS, Col)
                If iX > 0 Then
                    X = .Cells(i, iX)
                    Rqst = FF.Cells(PTRN_ADAPT, Col)
                    F_rqst = FF.Cells(PTRN_FETCH, Col)

                    Y = Adapter(Rqst, X, F_rqst, IsErr, R.EOL, i, Col)

                    If Not IsErr Then .Cells(i, Col) = Y
                ElseIf iX < 0 Then
                    Exit For
                End If
            Next Col
        Next i
    End With
End Sub
Function Adapter(Request, ByVal X, F_rqst, IsErr, Optional EOL_Doc, Optional iRow, Optional iCol) As String
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

    Dim FF() As String, Tmp() As String
    Dim i As Long, Par() As String
    Dim WP_Row As Long  ' ������ ��� ������ ��������� ���������, ����������� � Select
    
    IsErr = False
    X = Trim(X)
    
'--- ������ ������ �������� ���� <���>/C1,C2,C3...
    Dim AdapterName As String
    AdapterName = ""
    If Request <> "" Then
        Tmp = Split(Request, "/")
        AdapterName = Tmp(0)
        If InStr(Request, "/") <> 0 Then Par = Split(Tmp(1), ",")
    End If

'======== ������������� �������� ��� �������������� ��������� X ����� Fetch =========
    Select Case AdapterName
    Case "MainContract":
        X = Trim(Replace(X, "�������", ""))
    Case "<>0":
        If X = "0" Then X = ""
    Case "ContrK":
        Const PAY_REF = 8
        Dim MainDog As String, iPay As Long
        iPay = DB_TMP.Sheets(WP).Cells(PAY_REF, 4)
        MainDog = DB_1C.Sheets(PAY_SHEET).Cells(iPay, CLng(Par(0)))
        X = ContrCod(X, MainDog)
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
    If DB_TMP Is Nothing Then Set DB_TMP = FileOpen(F_TMP)
    With DB_TMP.Sheets(WP)
        Adapter = ""
        Select Case AdapterName
        Case "", "MainContract": Adapter = X
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
        Case "DogVal":
            Dim Vpaid As Long, Vinv As Long, Vdog As Long, DogCur As String
            Dim sDog As String
            Vpaid = .Cells(WP_PAYMENT_LINE, CLng(Par(0)))
            Vinv = .Cells(WP_PAYMENT_LINE, CLng(Par(1)))
            DogCur = .Cells(WP_PAYMENT_LINE, CLng(Par(2)))
            Vdog = 0
            sDog = Trim(.Cells(WP_PAYMENT_LINE, CLng(Par(3))))
            If sDog <> "" Then
                If Not IsNumeric(sDog) Then
                    ErrMsg FATAL_ERR, "�� �������� �������� � �������� ������� WP"
                    Stop
                    End
                End If
                Vdog = CDbl(sDog) * CurRate(DogCur)
            End If
            Adapter = Dec(Application.Max(Vpaid, Vinv, Vdog))
        Case "ForceTxt":
            Adapter = "'" & X
        Case "CopyToVal":
            WP_Row = iRow + .Cells(iRow + 3, 3) + PTRN_LNS - 1
            .Cells(iRow - 1 + PTRN_VALUE, iCol).Copy .Cells(WP_Row, iCol)
        Case "CopyFrVal":
            WP_Row = iRow + .Cells(iRow + 3, 3) + PTRN_LNS - 1
            .Cells(WP_Row, iCol).Copy .Cells(iRow - 1 + PTRN_VALUE, iCol)
        Case "OppFilter":
            Const SEL_REF = 20
        ' ��������� ���� �� ������ ��������� � ���������
            Dim IdSFopp As String
            IdSFopp = .Cells(SEL_REF, 3)
            If IdSFopp = "" Then
                Dim b As Long, A(0 To 6) As Long
                b = .Cells(SEL_REF + 2, 4)
                For i = 0 To UBound(A)
                    A(i) = CLng(Par(i))
                Next i
                Adapter = "-1"  ' -1 - �������, ��� ��������� EOL, � ������ �� ������
                For i = .Cells(SEL_REF, 4) + 1 To EOL_Doc
                    If OppFilter(i, .Cells(b, A(0)), .Cells(b, A(1)), _
                            .Cells(b, A(2)), .Cells(b, A(3)), .Cells(b, A(4)), _
                            .Cells(b, A(5)), .Cells(b, A(6))) Then
                        Adapter = i
                        Exit For
                    End If
                Next i
            Else
    ' ������� ���� ������������ ������, ����� ������ � ���������, � �� ������ � ��������
                Dim Rdoc As TOCmatch, Doc As String
                Doc = .Cells(iRow, 1)
                Rdoc = GetRep(Doc)
                Adapter = CSmatchSht(X, SFOPP_OPPID_COL, Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN))
                .Cells(iRow + PTRN_LNS, 11) = "�������"
                .Cells(iRow + PTRN_LNS, 11).Interior.Color = rgbBlue
                If Adapter = .Cells(iRow + 1, 4) Then Adapter = "-1"
            End If
        Case "SetOppButton":
            WP_Row = iRow + .Cells(iRow + 3, 3) + PTRN_LNS - 1  ' ����������� ������ "�������"
            .Cells(iRow - 1 + PTRN_VALUE, iCol).Copy .Cells(WP_Row, iCol)
            If X = "" Then Adapter = "�������"  ' ���� � ������� ��� �������� - ������ "�������"
        Case "NewOppName":
    ' -- ��������� ��� ������� � ���� �����������-��������� ������� ����
            Dim Typ As String, Dogovor As String, Dat As String
            Typ = .Cells(WP_PAYMENT_LINE, CLng(Par(0)))
            Dogovor = .Cells(WP_PAYMENT_LINE, CLng(Par(1)))
            MainDog = .Cells(WP_PAYMENT_LINE, CLng(Par(2)))
            Dogovor = ContrCod(Dogovor, MainDog)
            Dat = .Cells(WP_PAYMENT_LINE, CLng(Par(3)))
            Adapter = X & "-" & Typ & " " & Dogovor & " " & Dat
        Case "TypOpp":
    ' -- ������������� ���� ������� �� ���� � ������������ ������
            Dim good As String
'            Stop
            good = .Cells(WP_PAYMENT_LINE, CLng(Par(0)))
            Adapter = TypOpp(X, good)
        Case Else
            ErrMsg FATAL_ERR, "Adapter> �� ���������� " & AdapterName
        End Select
    End With
    Exit Function
    
AdapterFailure:
    ErrMsg WARNING, "������� " & AdapterName & "(" & X & ") �� ������� ������"
    IsErr = True
End Function
Function X_Parse(iRow, iCol, _
    Optional PutToRow, Optional PutToCol, Optional iLine) As String
'
' -  X_Parse(iRow, iCol [, PutToRow, PutToCol, iLine])  - ������ ������ � - ��������� ��������
'   ����� (iRow,iCol)       - ����� ������ ������� ��� �������, ������ �� ����� �������
'       [PutToRow,PutToCol] - ����� ������, ���� ��������� ��������� ��������
'           [iLine]         - ����� ������ ��� ������� ���� iLine
'
' � ���� ������� �������� ����������� #6 ��� !6
'  * ���� # ��������, ��� ���������� �� ������� � ActiveSheet, � ������� ������ �������
'  * ���� ! - ��� � ���� ������� �� ������ WProw ���� �������
'
' 22.10.12
' 25.10.12 - ���������� � ����� � HashFlag=True
' 11.11.12 - �������� ��������� !<Col> ��� ��������� WProw
' 15.11.12 - Optional ���������

    Dim X_rqst As String, sX() As String
    Dim PtrnType As String
    Dim iX As Long, WP_Row As Long
    Dim RefType As String
    
    X_Parse = ""
    
    With DB_TMP.Sheets(WP)
        WP_Row = iRow - 1 + PTRN_VALUE
        
        PtrnType = .Cells(iRow, 2)
        If PtrnType = PTRN_SELECT Then WP_Row = iRow + PTRN_LNS + .Cells(iRow + 3, 3) - 1
        
        PutToRow = WP_Row: PutToCol = iCol
        
        X_rqst = .Cells(iRow - 1 + PTRN_COLS, iCol)
        
        If X_rqst = "" Then GoTo Ex
        sX = Split(X_rqst, "/")
        
        RefType = Left(sX(0), 1)
        If RefType = "#" Or RefType = "!" Then sX(0) = Mid(sX(0), 2)
        
        iX = 0
        If UBound(sX) >= 0 Then iX = sX(0)
        If iX > 0 Then
            Select Case PtrnType
            Case "������", "������": GoTo GetFromWP
            Case "iLine":
                WP_Row = iLine
                GoTo GetFromActiveSheet
            Case PTRN_SELECT:
                WP_Row = .Cells(PutToRow, 5)
                GoTo GetFromActiveSheet
             Case Else:
                ErrMsg FATAL_ERR, "xAdapt> �������� ��� ������� " & PtrnType
            End Select
        End If

GetFromWP:
        If iX > 0 Then X_Parse = .Cells(WP_Row, iX)
        GoTo Ex
    End With
    
GetFromActiveSheet:
    If RefType = "!" Then
        WP_Row = PutToRow
        GoTo GetFromWP
    ElseIf RefType = "#" Then
        WP_Row = iRow + PTRN_VALUE - 1
        GoTo GetFromWP
    End If
    If iX > 0 Then X_Parse = ActiveSheet.Cells(WP_Row, iX)
Ex: Exit Function
End Function
Function FetchDoc(F_rqst, X, IsErr) As String
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
' *             ������ ������ - ��������� ��������� ������ Fetch - /W ��� /0
' *  /W             - WARNING � Log, ��������� IsErr=False, ���� ��������� ""
' *  /0             - "" ������ ��������� (��������, ������� � ������)
' *  /D             - Default - "" ��������, �� IsErr=True ��� �������� �� ���������
' *             ������ ������ ����������� - �������� Log � IsErr = True
'
' 5.9.12
' 14.9.12 - �������� /D ��� ������ ������ - "�� ���������"
' 4.11.12 - Fetch ���������� ����� ������ � ������ <Doc>/C1:�

    FetchDoc = ""
    If F_rqst = "" Or X = "" Then GoTo ErrExit
        
    Dim Tmp() As String, Cols() As String, s As String
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
' ������ Indx=� - ��� ������ �����, �� � ���������� ��� ���� Split
'!!!!!!!!!!!!!!!!!!!!!!!!!!!
        If Indx <= 0 Then
            ErrMsg WARNING, "FetchDoc: " & Doc & "(" & Indx & "," & C1 _
                & ") - ������������ ����� ������"
            GoTo ErrExit
        End If
        s = Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN).Cells(Indx, C1)
    Else
'--- �������� �1:C2 - � ������ 2 ��������� - ��������� �������� �� Lookup ��� �
        If IsNumeric(Cols(1)) Then C2 = Cols(1)
        s = ""
        N = CSmatchSht(X, C1, Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN))
        If N <> 0 Then
            If Cols(1) = "�" Then
                s = N
            Else
                s = Workbooks(Rdoc.RepFile).Sheets(Rdoc.SheetN).Cells(N, C2)
            End If
        End If
    End If
'--- ��������� ������ 2 -- ���� S=""
    If s = "" Then
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
        FetchDoc = s
    End If
    
OK_Exit:    IsErr = False
    Exit Function
ErrExit:    IsErr = True

End Function
