Attribute VB_Name = "NewEntities"
'-----------------------------------------------------------------------------
' NewEntities   - ����� �������, ��������, etc � "�������" ������ WP_TMP
'
' S NewSheet(SheetName, HDRform) - ������� ����� ���� SheetName
'       �������� ����� ������ ����� ������� �� �������� SheetName,
'       � ������ ������� �����- �� ������� c����� �����

'   27.10.2012

Option Explicit

Sub NewSheet(SheetName As String)
'
' S NewSheet(SheetName, HDRform) - ������� ����� ���� SheetName
'       �������� ����� ������ ����� ������� �� �������� SheetName,
'       � ������ ������� �����- �� ������� c����� �����
'
' 19.8.12
'  3.9.12 - StepIn
'  9.9.12 - displayAlert = False ��� Delete Sheet
'  1.10.12 - bug fix
' 19.10.12 - ����������� "�������" ������ � WP_TMP
' 27.10.12 - ������� "�������" ������ � ����� ������� TOCmatch

    StepIn
    
    Dim R As TOCmatch
    Dim i As Long, Cols As Long, W As String
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
            .Tab.Color = rgbLightBlue
            For i = 1 To Cols
                Frm.Columns(i).Copy Destination:=.Cells(1, i)
                W = .Cells(3, i)
                If IsNumeric(W) Then .Cells.Columns(i).ColumnWidth = CDbl(W)
            Next i
            For i = 2 To .UsedRange.Rows.Count
                .Rows(2).Delete
            Next i
        End With
    End With
'-- ���������� � TOCmatch ������ �� ������ �����
    R.EOL = EOL(R.SheetN, DB_TMP)
    If R.EOL <> 1 Then GoTo ErrHdr
    R.CreateDat = Now
    WrTOC           ' ��������� ���� � TOCmatch ������� StepOut
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
Sub NewPay(i, OppN, ContrId)
'
' ����� ������ � SF �� ������ i 1� - ������� ����������� DL
'   21.2.12
'   28.2.12 - ��� EOL_NewPay

    Dim j  ' ��������� ������ ����� P_Paid
    
    j = EOL_NewPay + 1
    EOL_NewPay = j
    
    With DB_MATCH.Sheets(P_Paid)
        .Cells(j, 1) = DB_1C.Sheets(PAY_SHEET).Cells(i, 6)            ' ����.���.
        .Cells(j, 2) = DDMMYYYY(DB_1C.Sheets(PAY_SHEET).Cells(i, 7))  ' ����
        .Cells(j, 3) = DB_1C.Sheets(PAY_SHEET).Cells(i, 8)            ' ����
        .Cells(j, 4) = Dec(DB_1C.Sheets(PAY_SHEET).Cells(i, 18))      ' ���� ���.
        .Cells(j, 5) = DB_1C.Sheets(PAY_SHEET).Cells(i, 19)           ' �����
        .Cells(j, 6) = ContrId                          ' ContractId
        .Cells(j, 7) = OppN                             ' OppN
    End With
End Sub
Sub NewOpp(Account, ContrK, CloseDate, Sale, Value, CurrencyOpp, TypGood, Sbs, _
    Optional Stage = "90%-������ ������ ������ �� ����")
'
' ����� ������ ��� ������ DL � ����������� Account.
'   12.2.2012
'   12.3.12 - ���������� ���� ��������� = OppBuddy
'   13.3.12 - ����������� ������ NewOpp
'   23.3.12 - ����������� ��������� ����� Public ExRespond=False
'   22.4.12 - bug fix � Dec(Value)
'   24.4.12 - ������ ������ � Line, Kind, OppType

    Dim i, AccountId, OwId, OwnerOpp, N, V
    Dim OppBuddy, OppName, OppUniq As String
        
    AccountId = AccId(Account)
    OwId = OwnerId(Sale, OppBuddy)
    OppUniq = Account & "-" & TypGood & " " & ContrK
    OppName = OppUniq
'---- ������ ��� �������� ������ ������� ----------------
    If TypGood <> "����������" Then
        OppName = OppUniq & " " & CloseDate
        If Value < MinNewOpp Then Exit Sub
        If Value < MinNewOppDialog Then
            Dim R
            R = MsgBox("������ '" & OppName & "'" & vbCrLf & vbCrLf _
                & Value & " " & CurrencyOpp _
                & " -- �������? ", vbYesNoCancel)
            If R <> vbYes Then
                If R = vbCancel Then ExRespond = False
                Exit Sub
            End If
        End If
    Else
'---- ������������ �������� �� �����������:
'           � ����������� ���������� ���� ��������� ������ ��� �����������
        With DB_SFDC.Sheets(SFopp)
            For i = 1 To EOL_SFopp
                If .Cells(i, SFOPP_ACC1C_COL) = Account _
                        And .Cells(i, SFOPP_CLOSEDATE_COL) >= DATE_BULKY _
                    Then Exit Sub
            Next i
        End With
    End If
    
    With DB_MATCH.Sheets(O_NewOpp)
'---- ������������ ����� ����������� �������� -----------
        For i = 1 To EOL_NewOpp
            If .Cells(i, 1) = AccountId Then    ' ���� �� ��� ������� �������� NewOpp
            
                If .Cells(i, NEWOPP_OPPUNIQ_COL) = OppUniq Then
                    If TypGood = "����������" Then
                        Exit Sub
                    Else
                        V = WorksheetFunction.Substitute(.Cells(i, NEWOPP_VALUE_COL), ".", ",")
                        .Cells(i, NEWOPP_VALUE_COL) = Dec(Value + V)
                    End If
                    Exit For
                End If
            End If
        Next i

' ----- ���������� ����������� ����., ���� � ��� ������� �� TypGood ----
        Dim Line, Kind, OppType As String
        Dim iG As Range
        For Each iG In Range("GoodSbs").Rows
            If iG.Cells(1, WE_GOODS_COL) = TypGood Then
                Line = iG.Cells(1, WE_GOODS_LINCOL)
                Kind = iG.Cells(1, WE_GOODS_KINDCOL)
                OppType = TypGood                       ' ������������� ���� �������
                If TypGood = "������������" Then OppType = "������"
                If iG.Cells(1, WE_GOODS_ISSBSCOL) <> "" Then
                    OppType = iG.Cells(1, WE_GOODS_ISSBSCOL)
                    If Not Sbs Then OppType = iG.Cells(1, WE_GOODS_NOSBSCOL)
                End If
                Exit For
            End If
        Next iG
       
' ----- ������� ������ ������� � NewOpp, ����������� ����� -------------
        EOL_NewOpp = EOL_NewOpp + 1
        N = EOL_NewOpp
        .Cells(N, NEWOPP_ACCID_COL) = AccountId
        .Cells(N, NEWOPP_OPPNAME_COL) = OppName
        .Cells(N, NEWOPP_CLOSDATE_COL) = DDMMYYYY(CloseDate)
        .Cells(N, NEWOPP_OWNERID_COL) = OwId
        .Cells(N, NEWOPP_VALUE_COL) = Dec(Value)
        .Cells(N, NEWOPP_ISOCUR_COL) = CurrencyOpp
        .Cells(N, NEWOPP_TYPOPP_COL) = "��������"
        .Cells(N, NEWOPP_STAGENAME_COL) = Stage
        .Cells(N, NEWOPP_TYPOPP_COL) = OppType
        .Cells(N, NEWOPP_LINE_COL) = Line
        .Cells(N, NEWOPP_KIND_COL) = Kind
        .Cells(N, NEWOPP_BUDDY_COL) = OppBuddy          ' ���������
        .Cells(N, NEWOPP_OPPUNIQ_COL) = OppUniq         ' ��������� ���� OppUniq
        If TypGood = "����������" Then
            .Cells(N, NEWOPP_CLOSDATE_COL) = "1.1.2020"
            .Cells(N, NEWOPP_VALUE_COL) = Dec(999999)
        End If
    End With
End Sub


