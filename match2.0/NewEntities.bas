Attribute VB_Name = "NewEntities"
'-----------------------------------------------------------------------------
' NewEntities   - ����� �������, ��������, etc � "�������" ������ WP_TMP
'
' S NewSheet(SheetName, TabColor) - ������� ����� ���� SheetName
'       �������� ����� ������ ����� ������� �� �������� SheetName,
'       � ������ ������� �����- �� ������� c����� �����
' S NewOrder(NewOrder)  - �������� ������� ��� ��������� � SF ����� ����� DL
'   7.5.2013

Option Explicit

Sub testNewSheet()
a:
    Set DB_MATCH = FileOpen(F_MATCH)
    DB_MATCH.Sheets("Process").Cells(1, PROCESS_NAME_COL) = "HANDL_PaidOpp"
    DB_MATCH.Sheets("Process").Cells(1, STEP_NAME_COL) = "NewSheet"
    NewSheet "NewPayment"
'    WrNewSheet "NewPayment", F_TMP, 3
    Stop
    GoTo a
End Sub

Sub NewSheet(SheetName As String, Optional TabColor As Long = rgbLightBlue)
'
' S NewSheet(SheetName, TabColor) - ������� ����� ���� SheetName
'       �������� ����� ������ ����� ������� �� �������� SheetName,
'       � ������ ������� �����- �� ������� c����� �����
'
' 19.8.12
'  3.9.12 - StepIn
'  9.9.12 - displayAlert = False ��� Delete Sheet
'  1.10.12 - bug fix
' 19.10.12 - ����������� "�������" ������ � WP_TMP
' 27.10.12 - ������� "�������" ������ � ����� ������� TOCmatch
' 23.11.12 - Optional TabColor
' 16.01.13 - ������������� setColWidth, ������� ������ �������
' 28.01.13 - width � setColWidth ������ ������: ������/������

    StepIn
    
    Dim R As TOCmatch
    Dim i As Long, Cols As Long
    Dim Frm As Range
    
    R = GetRep(SheetName)
    
    On Error GoTo NoHdr
    Set Frm = DB_MATCH.Sheets(Header).Range(R.FormName)
    Cols = Frm.Columns.Count
    On Error GoTo 0
    
    If DB_TMP Is Nothing Then Set DB_TMP = FileOpen(F_TMP)
    With DB_TMP
'-- ���������� ������� ����������� ����
        Application.DisplayAlerts = False
        On Error Resume Next
        .Sheets(SheetName).Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
        
        .Sheets.Add After:=.Sheets(.Sheets.Count)
        .Sheets(.Sheets.Count).Name = SheetName
        With .Sheets(SheetName)
            .Tab.Color = TabColor
            For i = 1 To Cols
                Frm.Columns(i).Copy Destination:=.Cells(1, i)
'                If IsNumeric(W) Then .Cells.Columns(i).ColumnWidth = CDbl(W)
                setColWidth DB_TMP.Name, SheetName, i, .Cells(3, i)
                
            Next i
            For i = 2 To .UsedRange.Rows.Count
                .Rows(2).Delete
            Next i
        End With
    End With
'-- ���������� � TOCmatch ������ �� ������ �����
    R.EOL = EOL(R.SheetN, DB_TMP)
    If R.EOL <> 1 Then GoTo ErrHdr
    R.CreateDat = Now        ' ��������� ���� � TOCmatch ������� StepOut
    RepTOC = R
    WrTOC
    Exit Sub
NoHdr:
    ErrMsg FATAL_ERR, "NewSheet> ��� ������� (�����) '" & R.FormName _
        & "' ��� ����� " & SheetName
    End
ErrHdr:
    ErrMsg FATAL_ERR, "NewSheet> ������ ������� (�����) '" & R.FormName _
        & "' ��� ����� " & SheetName & " -- ������������ EOL"
    End
End Sub
'''Sub NewPay(i, OppN, ContrId)
''''
'''' ����� ������ � SF �� ������ i 1� - ������� ����������� DL
''''   21.2.12
''''   28.2.12 - ��� EOL_NewPay
'''
'''    Dim j  ' ��������� ������ ����� P_Paid
'''
'''    j = EOL_NewPay + 1
'''    EOL_NewPay = j
'''
'''    With DB_MATCH.Sheets(P_Paid)
'''        .Cells(j, 1) = DB_1C.Sheets(PAY_SHEET).Cells(i, 6)            ' ����.���.
'''        .Cells(j, 2) = DDMMYYYY(DB_1C.Sheets(PAY_SHEET).Cells(i, 7))  ' ����
'''        .Cells(j, 3) = DB_1C.Sheets(PAY_SHEET).Cells(i, 8)            ' ����
'''        .Cells(j, 4) = Dec(DB_1C.Sheets(PAY_SHEET).Cells(i, 18))      ' ���� ���.
'''        .Cells(j, 5) = DB_1C.Sheets(PAY_SHEET).Cells(i, 19)           ' �����
'''        .Cells(j, 6) = ContrId                          ' ContractId
'''        .Cells(j, 7) = OppN                             ' OppN
'''    End With
'''End Sub
''''Sub NEWOPP(Account, ContrK, CloseDate, Sale, Value, CurrencyOpp, TypGood, Sbs, _
''''    Optional Stage = "90%-������ ������ ������ �� ����")
'''''
''''' ����� ������ ��� ������ DL � ����������� Account.
'''''   12.2.2012
'''''   12.3.12 - ���������� ���� ��������� = OppBuddy
'''''   13.3.12 - ����������� ������ NewOpp
'''''   23.3.12 - ����������� ��������� ����� Public ExRespond=False
'''''   22.4.12 - bug fix � Dec(Value)
'''''   24.4.12 - ������ ������ � Line, Kind, OppType
''''
''''    Dim i, AccountId, OwId, OwnerOpp, N, V
''''    Dim OppBuddy, OppName, OppUniq As String
''''
''''    AccountId = AccId(Account)
''''    OwId = OwnerId(Sale, OppBuddy)
''''    OppUniq = Account & "-" & TypGood & " " & ContrK
''''    OppName = OppUniq
'''''---- ������ ��� �������� ������ ������� ----------------
''''    If TypGood <> "����������" Then
''''        OppName = OppUniq & " " & CloseDate
''''        If Value < MinNewOpp Then Exit Sub
''''        If Value < MinNewOppDialog Then
''''            Dim R
''''            R = MsgBox("������ '" & OppName & "'" & vbCrLf & vbCrLf _
''''                & Value & " " & CurrencyOpp _
''''                & " -- �������? ", vbYesNoCancel)
''''            If R <> vbYes Then
''''                If R = vbCancel Then ExRespond = False
''''                Exit Sub
''''            End If
''''        End If
''''    Else
'''''---- ������������ �������� �� �����������:
'''''           � ����������� ���������� ���� ��������� ������ ��� �����������
''''        With DB_SFDC.Sheets(SFopp)
''''            For i = 1 To EOL_SFopp
''''                If .Cells(i, SFOPP_ACC1C_COL) = Account _
''''                        And .Cells(i, SFOPP_CLOSEDATE_COL) >= DATE_BULKY _
''''                    Then Exit Sub
''''            Next i
''''        End With
''''    End If
''''
''''    With DB_MATCH.Sheets(O_NewOpp)
'''''---- ������������ ����� ����������� �������� -----------
''''        For i = 1 To EOL_NewOpp
''''            If .Cells(i, 1) = AccountId Then    ' ���� �� ��� ������� �������� NewOpp
''''
''''                If .Cells(i, NEWOPP_OPPUNIQ_COL) = OppUniq Then
''''                    If TypGood = "����������" Then
''''                        Exit Sub
''''                    Else
''''                        V = WorksheetFunction.Substitute(.Cells(i, NEWOPP_VALUE_COL), ".", ",")
''''                        .Cells(i, NEWOPP_VALUE_COL) = Dec(Value + V)
''''                    End If
''''                    Exit For
''''                End If
''''            End If
''''        Next i
''''
''''' ----- ���������� ����������� ����., ���� � ��� ������� �� TypGood ----
''''        Dim Line, Kind, OppType As String
''''        Dim iG As Range
''''        For Each iG In Range("GoodSbs").Rows
''''            If iG.Cells(1, WE_GOODS_COL) = TypGood Then
''''                Line = iG.Cells(1, WE_GOODS_LINCOL)
''''                Kind = iG.Cells(1, WE_GOODS_KINDCOL)
''''                OppType = TypGood                       ' ������������� ���� �������
''''                If TypGood = "������������" Then OppType = "������"
''''                If iG.Cells(1, WE_GOODS_ISSBSCOL) <> "" Then
''''                    OppType = iG.Cells(1, WE_GOODS_ISSBSCOL)
''''                    If Not Sbs Then OppType = iG.Cells(1, WE_GOODS_NOSBSCOL)
''''                End If
''''                Exit For
''''            End If
''''        Next iG
''''
''''' ----- ������� ������ ������� � NewOpp, ����������� ����� -------------
''''        EOL_NewOpp = EOL_NewOpp + 1
''''        N = EOL_NewOpp
''''        .Cells(N, NEWOPP_ACCID_COL) = AccountId
''''        .Cells(N, NEWOPP_OPPNAME_COL) = OppName
''''        .Cells(N, NEWOPP_CLOSDATE_COL) = DDMMYYYY(CloseDate)
''''        .Cells(N, NEWOPP_OWNERID_COL) = OwId
''''        .Cells(N, NEWOPP_VALUE_COL) = Dec(Value)
''''        .Cells(N, NEWOPP_ISOCUR_COL) = CurrencyOpp
''''        .Cells(N, NEWOPP_TYPOPP_COL) = "��������"
''''        .Cells(N, NEWOPP_STAGENAME_COL) = Stage
''''        .Cells(N, NEWOPP_TYPOPP_COL) = OppType
''''        .Cells(N, NEWOPP_LINE_COL) = Line
''''        .Cells(N, NEWOPP_KIND_COL) = Kind
''''        .Cells(N, NEWOPP_BUDDY_COL) = OppBuddy          ' ���������
''''        .Cells(N, NEWOPP_OPPUNIQ_COL) = OppUniq         ' ��������� ���� OppUniq
''''        If TypGood = "����������" Then
''''            .Cells(N, NEWOPP_CLOSDATE_COL) = "1.1.2020"
''''            .Cells(N, NEWOPP_VALUE_COL) = Dec(999999)
''''        End If
''''    End With
''''End Sub

Sub UniqueHanle(NewSht As String, SFrep As String)
'
' S UniqueHanle(NewSht, SFrep)  - ���������� ���� Unique ����� NewSDht � ������������
'
' �������� �� ���� ������� NewSht:
'  1. ���� ����� ��� ������� � SF - NOP
End Sub
Sub NewOrder(NewOrd As String)
'
' S NewOrder(NewOrder)  - �������� ������� ��� ��������� � SF ����� ����� DL
'
' �������� �� ���� �������:
'  1. ���� ����� ��� ������� � SF - NOP
'  2. ���� ����� ��� ���� ����� CSD ��������������� - NOP
'  3. ���� ������
'       - ���� �� ��������, ��� � ������ -!!- ����� RightSale
'       - � ��� �� ������� ����� 1� -!!- ����� ��������� � ��������� ������
'       - � ���������� ��� ������ "���������� ������"
'       - -!!- ����� ��������� ������������ ����������� - �������
' 26.4.2013
'  2.5.13 ����� � �������� Id ������� � SN � WrNewSheet ����� ������ ExtPar
'  7.5.13 ������������� FetchDoc ��� ���������� ���� "���������"

    StepIn
    
    Dim Ord As TOCmatch, P As TOCmatch
    Dim i As Long, j As Long, tmp
    Dim TMPsalesRep As String   ' �������� � ������
    Dim TMPinv1C As String      ' ���� 1� � ������
    Dim TMPgoodType As String   ' ��������� ������ � ������
    Dim TMPcustomer As String   ' ������ � ������
    Dim CSDinvDate As Date      ' ���� ����� CSD
    Dim IdSFpaid As String      ' Id ������� � SF
    Dim ExtPar(3) As String     ' ������ ���������� ������������ � WrNewSheet
    Dim IdOpp As String, Team As String, IsErr As Boolean
    
    NewSheet NewOrd
    
    Ord = GetRep(ORDER_SHEET)
    P = GetRep(PAY_SHEET)

    With Workbooks(Ord.RepFile).Sheets(Ord.SheetN)
        For i = 2 To Ord.EOL
            Progress i / Ord.EOL
            If .Cells(i, OL_IDSF_COL) = "" Then
                tmp = .Cells(i, OL_CSDINVDAT_COL)
                If IsDate(tmp) Then
                    CSDinvDate = tmp
                Else
                    GoTo NextOrd
                End If
                If Trim(.Cells(i, OL_ORDERNUM_COL)) = "" Then GoTo NextOrd
                
                TMPsalesRep = Trim(.Cells(i, OL_SALES_COL))
                TMPinv1C = Trim(.Cells(i, OL_INV1C_COL))
                TMPgoodType = Trim(.Cells(i, OL_GOOD_COL))
                TMPcustomer = LCase$(.Cells(i, OL_CUSTOMER_COL))
                                
                For j = 2 To P.EOL
                    With Workbooks(P.RepFile).Sheets(P.SheetN)
                        If .Cells(j, PAYINSF_COL) <> 1 Then GoTo NextP
                        If TMPinv1C <> .Cells(j, PAYINV_COL) Then GoTo NextP
                        If Abs(.Cells(j, PAYDATE_COL) - CSDinvDate) > 50 Then GoTo NextP
                        If InStr(LCase$(.Cells(j, PAYACC_COL)), TMPcustomer) = 0 Then GoTo NextP
                        Select Case TMPgoodType
                        Case "ADSK": If .Cells(j, PAYGOODTYPE_COL) <> "Autodesk" Then GoTo NextP
                        Case "���� � ���": If .Cells(j, PAYGOODTYPE_COL) <> "����������" Then GoTo NextP
                        Case "NormaCS": If .Cells(j, PAYGOODTYPE_COL) <> "NormaCS" Then GoTo NextP
                        Case "SCAD", "������ ��":
                            If .Cells(j, PAYGOODTYPE_COL) <> "������ ��" Then GoTo NextP
                        Case "Altium": If .Cells(j, PAYGOODTYPE_COL) <> "Altium" Then GoTo NextP
                        Case "CS Dev": If .Cells(j, PAYGOODTYPE_COL) <> "�� CSoft" Then GoTo NextP
                        Case "Hard": If .Cells(j, PAYGOODTYPE_COL) <> "������������" Then GoTo NextP
                        Case Else
                        End Select
                        If TMPsalesRep = Trim(.Cells(j, PAYSALE_COL)) Then GoTo IdPfound
                        IdOpp = FetchDoc("SF/18:19", .Cells(j, PAYIDSF_COL), IsErr)
                        Team = FetchDoc("SFopp/1:11", IdOpp, IsErr)
                        If InStr(Team, TMPsalesRep) = 0 Then GoTo NextP
                        
IdPfound:               ExtPar(1) = .Cells(j, PAYIDSF_COL)  'Id ������� 1�
                        
                        WrNewSheet NewOrd, Ord.SheetN, i, ExtPar
                        Exit For
                    End With
NextP:          Next j
            End If
NextOrd: Next i
    End With
End Sub
