Attribute VB_Name = "PaidAnalitics"
'---------------------------------------------------------------------------
' PaidAnalitics -- ������� ��� ������� ���������� ��������
'
' S Paid1C()        - ������ �� ����� ��������, ��������� � SF � WP
' * PaidHandling()  - ������ �� ����� �������� 1�, ��������� � SF
' - GoodType(Good)              - ���������� ������ - ��� ������ Good
' - GoodJob(Good,GoodType,JobType)  - ���������� True ���� ����� Good ���� GoodType
'                                     ������������� ���� ����� JobType
' - IsSubscription(Good, GT)    - ���������� True, ���� ����� - ��������
'
'   15.9.2013

Option Explicit
Dim t0 As Single, t1 As Single, t2 As Single

Sub Paid1C(Optional ByVal iPayLine As Long = 2)
'
' S Paid1C ()    - ��������� �������� 1�, ��������� � SF � WP
'       * ���� ��� ������� �� �������� ���������� WP, ��� ����������� ������
'         WP_Adapt_Continue ��������� Run Paid1C(i)
'
' 9.9.13
' 12.9.13 - ����� DOG_UPDATE
' 15.9.13 - ������������� PAYCANONAME_COL ������ PAYACC_COL
'         - ������ �� ����� �� ������� �������� �� ������� ���������� �������

    StepIn
    
'           --- �������� ����� ---           ---�������---
    Const NEW_PAYMENT = "NewPayment":   Const HDR_WP = "HDR_WP"
    Const NEW_ACC = "NewAcc":           Const HDR_WPacc = "HDR_WPacc"
    Const NEW_OPP = "NewOpp":           Const HDR_WPopp = "HDR_WPopp"
    Const NEW_CONTRACT = "NewContract": Const HDR_WPcontract = "HDR_WPcontract"
    
    Const FETCH_ACC1C = Acc1C & "/" & A1C_NAME_COL & ":�/0"
'''    Const FETCH_SFD = "SFD/" & SFD_COD_COL & ":" & SFD_OPPID_COL & "/W"
    Const FETCH_SFD = "SFD/" & SFD_COD_COL & ":�/0"
    Const FETCH_SFOPP = "SFopp/" & SFOPP_ACC1C_COL & ":" & SFOPP_OPPID_COL & "/0"
    Const FETCH_DOGOVOR = DOG_SHEET & "/" & DOGCOD_COL & ":�/0"
    
    Const BALKY_TYPE = "����������"
    Dim LocalTOC As TOCmatch, i As Long, iLine As Long
    Dim sLine As String
    Dim IsErr As Boolean, FromN As Long
    Dim ContrK As String, OppId As String, ThisOppId As String
       
    LocalTOC = GetRep(PAY_SHEET)
    If iPayLine = 2 Then
        NewSheet NEW_PAYMENT
        NewSheet NEW_ACC
        NewSheet NEW_OPP
        NewSheet NEW_CONTRACT
        NewSheet DOG_UPDATE
    End If

    With Workbooks(LocalTOC.RepFile).Sheets(LocalTOC.SheetN)
        For i = iPayLine To LocalTOC.EOL
            Progress i / LocalTOC.EOL
            IsErr = False
            If .Cells(i, PAYINSF_COL) = 1 Then Exit For
'' 15.9.13                GoTo NextRow
            ElseIf .Cells(i, PAYISACC_COL) = "" Then
'!!'                WP_Adapt HDR_WPacc, i       '--- ���� ����������� ��� � SF
'!!'                iLine = FetchDoc(FETCH_ACC1C, .Cells(i, PAYCANONAME_COL), IsErr)
 ''''               If Not IsErr Then WrNewSheet NEW_ACC, Acc1C, iLine, HDR_NEWACC
                GoTo NextRow
            ElseIf Trim(.Cells(i, PAYDOGOVOR_COL)) <> "" Then   ' ���� ������� � 1�
                ContrK = ContrCod(.Cells(i, PAYDOGOVOR_COL), .Cells(i, PAYOSNDOGOVOR_COL))
                sLine = FetchDoc(FETCH_SFD, ContrK, IsErr)
                If sLine = "" Or Not IsNumeric(sLine) Then
                                            '--- ���� ������� �� ������� � SF
                    sLine = FetchDoc(FETCH_DOGOVOR, ContrK, IsErr)
                    If sLine = "" Or Not IsNumeric(sLine) Then
                        ErrMsg WARNING, "� ������ " & i & "�������� 1� ������ �� ������� " _
                            & ContrK & ", �������� ��� � ����� ��������� 1�"
                        GoTo NextRow
                    End If
                    WrNewSheet NEW_CONTRACT, DOG_SHEET, CLng(sLine)
                    GoTo NextRow
                End If
                OppId = DB_SFDC.Sheets(SFD).Cells(CLng(sLine), SFD_OPPID_COL)
                If OppId = "" Then
'--- ���� ������ � ����� �������� � ��������
                    Dim Opps() As Long
                    Opps = OppSelect(i)
                    If Opps(0) = 1 Then
                        Dim S As String
                        S = DB_SFDC.Sheets(SFopp).Cells(Opps(1), SFOPP_OPPID_COL)
                        WrNewSheet DOG_UPDATE, PAY_SHEET, iPayLine, S
                    Else
                        WP_Adapt HDR_WP, i   '--- ���� � SF ��� ������� �� ��������
                    End If
                    GoTo NextRow
                End If
            ElseIf .Cells(i, PAYGOODTYPE_COL) = BALKY_TYPE Then
                Dim BalkyExists As Boolean: BalkyExists = False
                FromN = 2                   '--- ��������� ���������� ������ Balky
                Do While FromN <> 0
                    ThisOppId = FetchDoc(FETCH_SFOPP, .Cells(i, SFOPP_ACC1C_COL), IsErr, FromN)
                    If ThisOppId = "" Then GoTo NextOpp
                    With DB_SFDC.Sheets(SFopp)
                        If .Cells(FromN, SFOPP_LINE_COL) <> BALKY_TYPE Then GoTo NextOpp
                        If .Cells(FromN, SFOPP_CLOSEDATE_COL) - Now < 365 Then GoTo NextOpp
                    End With
                    If BalkyExists Then
                        ErrMsg WARNING, "� ����������� '" & .Cells(i, PAYCANONAME_COL) & "' ��������� �������� �� �����������"
                        GoTo NextRow
                    End If
                    OppId = ThisOppId
NextOpp:            FromN = FromN + 1
                Loop
                GoTo ToSF
            Else
                WP_Adapt HDR_WP, i
                GoTo NextRow
            End If
ToSF:       If Not IsErr Then WrNewSheet NEW_PAYMENT, PAY_SHEET, i, OppId
NextRow:
        Next i
    End With
    SheetDedup NEW_CONTRACT, 1
End Sub
Sub TestOppSelect()
'
' T ������� OppSelect
'       11.9.13

    Dim N() As Long, TT As Long
''''    TT = 0
''''    TT = UBound(N)
''''    On Error Resume Next
''''    TT = UBound(N) - LBound(N)
''''    On Error GoTo 0
    N = OppSelect(2)
    N = OppSelect(1707)
End Sub

Function OppSelect(ByVal iPaid As Long) As Long()
'
' - OppSelect(Account, Sale)    - ������� ������ ����� �������� �� SFopp �� ����� WP
'                                 �� ��������  iPaid - ������ ������ �������.
'                                 ���������� ������ ������� ����� SFopp ��������� ��������.
' 6.9.13
' 11.9.13 - ���������� ������ �������� - ������� ����� SFopp,
'           ������ � �������� (0) ����� ��������� ��������
' 15.9.13 ������������� PAYCANONAME_COL ������ PAYACC_COL

    Const FETCH_SFOPP = "SFopp/" & SFOPP_ACC1C_COL & ":�"
    
    Dim Opps() As Long, nOpp As Long, maxNopp As Long
    Dim LocalTOC As TOCmatch, iOpp As Long, IsErr As Boolean
    Dim sN As String
    Dim Account As String, Salesman As String, PaidDate As Date
    Dim ContrK As String, Rub As Long, GoodT As String

    GetRep PAY_SHEET
    With DB_1C.Sheets(PAY_SHEET)
        Account = .Cells(iPaid, PAYCANONAME_COL)
        Salesman = .Cells(iPaid, PAYSALE_COL)
        ContrK = ContrCod(.Cells(iPaid, PAYDOGOVOR_COL), .Cells(iPaid, PAYOSNDOGOVOR_COL))
        PaidDate = .Cells(iPaid, PAYDATE_COL)
        Rub = .Cells(iPaid, PAYRUB_COL)
'''        GoodT = .Cells(iPaid, PAYGOODTYPE_COL)
        GoodT = GoodType(.Cells(iPaid, PAYGOOD_COL))
    End With
    
    LocalTOC = GetRep(SFopp)
    With Workbooks(LocalTOC.RepFile).Sheets(LocalTOC.SheetN)
        iOpp = 1: maxNopp = 0
        ReDim Opps(0): Opps(0) = 0
        Do
            sN = FetchDoc(FETCH_SFOPP, Account, IsErr, iOpp + 1)
            If IsErr Or Not IsNumeric(sN) Then Exit Do
            iOpp = sN
            sN = .Cells(iOpp, SFOPP_OPPN_COL): nOpp = sN
            If Not IsSameTeam(Salesman, .Cells(iOpp, SFOPP_SALE_COL), nOpp) Then GoTo NxtOpp
'---            �� �������� � SFopp ��������� ��� ��� ������ ������� �������� ��� �������
'''            OppGoodType = .Cells(iOpp, SFOPP_TYP_COL)
'''            On Error Resume Next
'''            OppGoodType = WorksheetFunction.VLookup(PayKod, _
'''                Sheets(SF).Range("B:S"), 18, False)
'''            On Error GoTo 0
'''
            If InStr(.Cells(iOpp, SFOPP_TYP_COL), GoodT) = 0 Then GoTo NxtOpp
        '-- ��������, ��� � ����� ������� ���� ����� �������� �� ������� 1�
            If InStr(.Cells(iOpp, SFOPP_OPPNAME_COL), ContrK) <> 0 Then GoTo FoundOppPaidDog
            maxNopp = maxNopp + 1
            ReDim Preserve Opps(maxNopp)
            Opps(maxNopp) = iOpp
''            OppSelect = OppSelect + 1
''            MsgBox "� ������� Sale=" & Salesman & " IsSameTeam=" & IsSameTeam(Salesman, .Cells(iOpp, SFOPP_SALE_COL), iOpp)
''            If OppFilter(i) Then
''                ����� ������ �� ������� WP
''            End If

''            End If
NxtOpp: Loop
    End With
    GoTo Ex
FoundOppPaidDog:    '���������� ������������ ������ � ������, ���������� ��� ��������
    ReDim Opps(1) As Long
    Opps(1) = iOpp: maxNopp = 1
Ex: Opps(0) = maxNopp   ' � Opps(0) ���������� ��������� ��������
    OppSelect = Opps
End Function
Function OppFilter(iOpp, Sale, Account, t, Rub, Dat, Dogovor, MainDog) As Boolean
'
' - OppFilter(iOpp, Sale, Account, T, Rub, Dat, Dogovor, MainDog) - ������� Select
'                   iOpp - ����� ������ � SFopp, ��������� ��������� �� �������
'                   ���������� False ���� ������ iSFopp �� ������������� ����������
' 21.10.12
' 31.10.12 - �������� ������ �� SalesTeam

    Dim ContrK As String, IdSFopp As String, iSFD As Long, OppN As Long
    OppFilter = False

    With DB_SFDC
        With .Sheets(SFopp)
            If .Cells(iOpp, SFOPP_ACC1C_COL) <> Account Then Exit Function
            OppN = .Cells(iOpp, SFOPP_OPPN_COL)
            If Not IsSameTeam(Sale, .Cells(iOpp, SFOPP_SALE_COL), OppN) Then Exit Function
        End With

        If DB_TMP.Sheets(WP).Cells(20, 3) = "" Then GoTo Found
        ContrK = ContrCod(Dogovor, MainDog)
        If ContrK = "" Then GoTo Found
        IdSFopp = .Sheets(SFopp).Cells(iOpp, SFOPP_OPPID_COL)
        iSFD = CSmatchSht(IdSFopp, SFD_CONTRID_COL, SFD)
        If iSFD > 0 Then
            If .Sheets(SFD).Cells(iSFD, SFD_COD_COL) = ContrK Then GoTo Found
        End If
    End With
    Exit Function


'''            OppN = .Cells(iOpp, SFOPP_OPPN_COL)
'''            OppT = .Cells(iOpp, SFOPP_TYP_COL)
'''            OppCur = .Cells(iOpp, SFOPP_TO_PAY_CUR_COL)
'''            OppToPayRub = .Cells(iOpp, SFOPP_TO_PAY_VAL_COL) * CurRate(OppCur)
'''            OppCloseDate = .Cells(iOpp, SFOPP_CLOSEDATE_COL)
'''            OppId = .Cells(iOpp, SFOPP_OPPID_COL)
'''            If InStr(OppT, SeekOppType) <> 0 _
'''                    And IsSameTeam(Sale, .Cells(iOpp, SFOPP_SALE_COL), OppN) _
'''                    And OppToPayRub >= Rub _
'''                    And Dat <= OppCloseDate Then
'''                If .Cells(iOpp, SFOPP_PROBABILITY_COL) <> 0 Then
'''                    GoTo Found
'''                Else
'''                    Msg = "� ����������� '" & Account & "'" _
'''                        & vbCrLf & vbCrLf & "���� ������ Closed/Lost" _
'''                        & vbCrLf & vbCrLf & OppName _
'''                        & vbCrLf & vbCrLf & "���������� ��� ���?"
'''                    Respond = MsgBox(Msg, vbYesNoCancel)
'''                    If Respond = vbCancel Then ExRespond = False
'''                    If Respond = vbYes Then
'''                        ErrMsg WARNING, "!! ��������� ��������� ������� " & OppN _
'''                            & vbCrLf & vbCrLf & "� ���� ������� ������!"
'''                        GoTo Found
'''                    End If
'''                End If
'''            End If
'''        End If

Found:
    OppFilter = True
End Function



'''Sub NewPaidOpp()
''''
'''' S NewPaidDog()    - ����� ������� �� ������������ ��������
''''   6.10.12
''''  31.10.12 - fix Call WrNewSheet
'''
'''    StepIn
'''
'''    Dim P As TOCmatch, S As TOCmatch
'''    Dim i As Long
'''
'''    Dim Dat As Date         '���� - "���� ����.�����"
'''    Dim Acc As String       '���� - "������"
'''    Dim Rub As Variant      '���� - "���� ���"
'''    Dim Sale As String      '���� - "��������"
'''    Dim good As String      '���� - "������" ����� ��������, ������������
'''    Dim T As String         ' = ��� ������ �� ������������
'''    Dim Dogovor As String   '���� - "�������"
'''    Dim MainDog As String   '���� - "���.�������"
'''    Dim ContrK As String    ' = ����� <���.�������>/<�������>
'''    Dim OppId As String     ' = Id ������� � SF
'''
'''    P = GetRep(PAY_SHEET)
'''    S = GetRep(SFopp)
'''    EOL_SFopp = S.EOL
'''    With DB_1C.Sheets(PAY_SHEET)
'''        For i = 2 To P.EOL
'''            Progress i / P.EOL
'''            If .Cells(i, PAYINSF_COL) <> 1 _
'''                    And Trim(.Cells(i, PAYISACC_COL)) <> "" _
'''                    And Trim(.Cells(i, PAYDOC_COL)) <> "" Then
'''                Acc = Compressor(.Cells(i, PAYCANONAME_COL)) ' �����������
'''                Dat = .Cells(i, PAYDATE_COL)    ' ���� �������
'''                Sale = .Cells(i, PAYSALE_COL)   ' ��������
'''                good = .Cells(i, PAYGOOD_COL)   ' �����
'''                T = GoodType(good)              ' ��� ������ �� �����
'''                Rub = .Cells(i, PAYRUB_COL)     ' ����� ������� ���
'''                Dogovor = .Cells(i, PAYDOGOVOR_COL)
'''                MainDog = Mid(.Cells(i, PAYOSNDOGOVOR_COL), 9)
'''                ContrK = ContrCod(Dogovor, MainDog)
'''
'''                OppId = IsOpp(Sale, Acc, T, Rub, Dat, ContrK) ' Id ������� � SF
'''
'''                If OppId <> "" Then
'''                    WrNewSheet NEW_PAYMENT, PAY_SHEET, i, OppId
'''                End If
'''            End If
'''        Next i
'''    End With
'''End Sub
'''Sub NewPaidDog()
''''
'''' S NewPaidDog()    - ����� ������� �� ��������
''''
'''' 31.10.21 fix Call WrNewSheet
'''
'''    StepIn
'''
'''    Dim P As TOCmatch
'''    Dim i As Long
'''
'''    P = GetRep(PAY_SHEET)
'''    With DB_1C.Sheets(PAY_SHEET)
'''        For i = 2 To P.EOL
'''            Progress i / P.EOL
'''            If .Cells(i, PAYISACC_COL) <> "" And .Cells(i, PAYINSF_COL) = "" Then
'''                If .Cells(i, PAYDOGOVOR_COL) <> "" Then
'''                    WrNewSheet NEW_PAYMENT, PAY_SHEET, i
'''                End If
'''            End If
'''        Next i
'''    End With
'''End Sub
'''Sub PaidHandling()
''''
'''' ������ �� �������� � ����� ����� �������� �� ���������� ��������
''''   16.8.12 match 2.0
'''
''''---- ���������� ��������� ���������� ----------
'''' � ������������ ���� "����' - �������� ���� �� ����� ��������1�
''''                        =   - ��������� ���������� � ���� ���������
'''    Dim PayK As String      '���� - ��� �������
'''    Dim Dat As Date         '���� - "���� ����.�����"
'''    Dim Acc As String       '���� - "������"
'''    Dim Rub As Variant      '���� - "���� ���"
'''    Dim Sale As String      '���� - "��������"
'''    Dim good As String      '���� - "������" ����� ��������, ������������
'''    Dim T As String         ' = ��� ������ �� ������������
'''    Dim Sbs As Boolean      ' = True ���� ������������ �������� ��������
'''    Dim Dogovor As String   '���� - "�������"
'''    Dim MainDog As String   '���� - "���.�������"
'''    Dim ContrK As String    ' = ����� <���.�������>/<�������>
'''    Dim ContrId As String   ' = Id �������� ��� � SF
'''    Dim OppId As String     ' = Id ������� � SF
'''
'''    Dim i                   ' = ������ ��������� ����� �� ��������
'''    Dim t0, Tbeg, TI    ' ��� �������������� �� �������
'''    t0 = Timer
''''-----------------------------------------------
'''
'''    Dim ts1 As Long, ts2 As Long, ts3 As Long, ts4 As Long, ts5 As Long ' profiling ?????????
'''    Dim ts1S As Long, ts2S As Long, ts3S As Long, ts4S As Long, ts5S As Long ' profiling ?????????
'''    ts1 = 0: ts2 = 0: ts3 = 0: ts4 = 0: ts2 = 0: ts5 = 0
'''
'''    TI = Timer
'''    LogWr t0 - TI & " PaidAnalitics: ������"
'''    Dim SumNewPay
'''    SumNewPay = 0
'''
'''
''''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'''    Dim P As TOCmatch, Opp As TOCmatch
'''    GetRep TOC
'''    DB_MATCH.Sheets(We).Activate
'''    ClearSheet O_NewOpp, Range("HDR_NewOpp")
'''    ClearSheet P_Paid, Range("HDR_Payment")
'''    Opp = GetRep(SFopp)
'''    EOL_SFopp = Opp.EOL
'''    P = GetRep(PAY_SHEET)
'''
'''    For i = 2 To P.EOL
'''        Progress (i / P.EOL)
'''        If ExRespond = False Then GoTo Ex
'''        With DB_1C.Sheets(P.SheetN)
''''''''            .Activate
'''        ' Account � SF ����, ������� � SF ���, ��� � �������� �� ������
'''            Acc = Compressor(.Cells(i, PAYCANONAME_COL)) ' �����������
'''            If .Cells(i, PAYISACC_COL) <> "" And _
'''                    Trim(.Cells(i, PAYDOC_COL)) <> "" And _
'''                    Trim(.Cells(i, PAYSALE_COL)) <> "" Then
'''                Dat = .Cells(i, PAYDATE_COL)    ' ���� �������
'''                Sale = .Cells(i, PAYSALE_COL)   ' ��������
'''                good = .Cells(i, PAYGOOD_COL)   ' �����
'''                T = GoodType(good)              ' ��� ������ �� �����
'''                Sbs = IsSubscription(good, T)   ' ��������?
'''                Rub = .Cells(i, PAYRUB_COL)     ' ����� ������� ���
'''                Dogovor = .Cells(i, PAYDOGOVOR_COL)
'''                MainDog = Mid(.Cells(i, PAYOSNDOGOVOR_COL), 9)
'''                ContrK = ContrCod(Dogovor, MainDog)
'''                ContrId = ContractId(ContrK)    ' Id ��������, ���� ����
'''                Dim IsInSF As String
'''                IsInSF = ""
'''
'''    '================ ������ ��� �� ������� � SF? =============================
'''                If .Cells(i, PAYINSF_COL) = "" Then
'''
'''                    ts1 = Timer                     ' tttttttttttttttttttttttttttttttttttttt
'''
'''                    OppId = IsOpp(Sale, Acc, T, Rub, Dat, ContrK) ' Id ������� � SF
'''                    If OppId = "" Then
'''                        NEWOPP Acc, ContrK, Dat, Sale, Rub, "RUB", T, Sbs
'''                    Else
'''            '>>>>  ��������� ������ �������
'''                        NewPay i, OppNbyId(OppId), ContrId
'''            '>>>>  ����� ������� � ��������� ��� �������� ������ �������� � SF
'''                        If ContrK <> "" And ContrId = "" Then
'''                            NewContract Dogovor, MainDog, ContrK
'''                        Else
'''                            ContrOppLink i, ContrK, ContrId, OppId
'''                        End If
'''                    End If
'''                    ts1S = ts1S + (Timer - ts1)      ' tttttttttttttttttttttttttttttttttttt
'''                    ts1 = ts1
'''                End If
'''            End If
'''        End With
'''    Next i
'''
''''    MsgBox "SumNewPay = " & SumNewPay
'''
'''    Dim ResultMsg
'''Ex: ResultMsg = "����: �������� " & EOL_NewPay - 1 & " ����� ��������; " _
'''        & EOL_NewOpp - 1 & " ����� ��������; " _
'''        & EOL_PaymentUpd - 1 & " �������� ������� � ����������; " _
'''        & EOL_ContrLnk - 1 & " ��������� ������� � ���������; " _
'''        & EOL_AdAcc - 1 & " ����� �����������;" _
'''        & EOL_ADSKlnkPay - 1 & " �������� ������� � ����������� Autodesk;"
'''    LogWr ResultMsg
'''    MsgBox ResultMsg
'''
'''    MsgBox "time1: " & ts1S & "; time2: " & ts2S & "; time3: " & ts3S & "; time4: " & ts4S
'''
'''    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Payment\"
'''    WriteCSV P_Paid, "Paid.txt"
'''    Shell "quota2.bat"
'''    WriteCSV P_PaymentUpd, "PmntUpd.txt"
'''    Shell "quota3.bat"
'''
'''    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\OppInsert\"
'''    WriteCSV O_NewOpp, "OppInsert.txt"
'''    Shell "quota2.bat"
'''
'''    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Dogovor\"
'''    WriteCSV NewContractLnk, "ContrUpd.txt"
'''    Shell "quota3.bat ContrUpd.txt C:\SFconstr\Dogovor\ContrUpd.csv"
'''
'''    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Account\"
'''    WriteCSV A_Acc, "AdAcc.txt"
'''    Shell "quota_Acc.bat"
'''
'''    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\ADSK\P_ADSK"
'''    WriteCSV P_ADSKlink, "P_ADSKlink.txt"
'''    Shell "quota_P_ADSK.bat"
'''
''''''    ModEnd 1
'''End Sub
'''Function IsOpp(Sale, Account, T, Rub, Dat, ContrCod)
''''
'''' ��������, ���� �� � ����������� Account �� ��������� ���������� ������ ���� �.
'''' ���� ������ �� �������� �� �����, ������� Rub - ���������� Id ����� �������.
'''' ��� �� ������������ ������� ���������� ������ ��� ����� ��� � ���������,
'''' ��������� � TargetContrK �� ������������ ������ ����� �������� � �������� �������.
'''' ���� ��� - ���������� ""
''''   13.2.2012
''''   21.2.2012 - ������� �� �������� ��������
''''   28.12.12 - ����� ����������� ������� ��� �������� �������
''''   3.3.12 - ��������� ���� �������� ������� OppCrDate,
''''   5.3.12 - Optional ContrCod ��� �������� ������� ����� ������ ������� � SF
''''   9.3.12 - �� ���� ������, ���� ������� �������� �� �������� ����� �
''''  16.3.12 - ������� ������� Closed/Lost, �� ���� ����������� = 0%
''''  18.3.12 - ���������� Id � �� ����� ���������� �������
''''  19.3.12 - TargetContrK - ������� ������� ��� ����� � �������� Opp
''''  22.3.12 - ������� ��� � � ����� ������� - ���� ������ - ��� ������������ Opp
''''  31.3.12 - ���������� ������ ������ ����������� �������
''''  20.4.12 - bug fix ��� �������� � ���������
''''  25.4.12 - ������ � ��������� �������������� � ��������� ������� �� ��������
''''  30.5.12 - ��������� � ��������� ������ "������������ �������� � �������"
''''  31.5.12 - ��������� �������� ������ � Close Lost ������
''''   5.10.12- ������ �� We � DB_MATCH
'''
'''    Dim i, OppToPayRub, OppCur, OppN, ContrOpp, OppName, OppId
'''    Dim OppCloseDate As Date, OppCrDat As Date
'''    Dim Msg As String           ' ��������� �� ������ � IsOpp
'''    Dim Respond As String       ' ����� � ������������� �������
'''
'''    IsOpp = ""
'''
'''    If T = "" Then Exit Function
'''    If Not IsRightSale(Sale, T) Then
'''        ErrMsg FATAL_ERR, "������������ �������� " & Sale & " � ������� " & T
'''        Exit Function
'''    End If
'''
'''    Dim SeekOppType As String, OppT As String
'''    SeekOppType = WorksheetFunction.VLookup(T, DB_MATCH.Sheets(We).Range("OppTypeRng"), 4, False)
'''
'''    With DB_SFDC.Sheets(SFopp)
''''-- ���� ������ ������ � ��������� -- ��������, ��� ����� � ContrCod ����
'''        If ContrCod <> "" Then
'''            For i = 2 To EOL_SFopp
'''                If .Cells(i, 4) = Account Then
'''                    OppName = .Cells(i, SFOPP_OPPNAME_COL)          ' ��� ��������
'''                    OppId = .Cells(i, SFOPP_OPPID_COL)
'''                    If InStr(OppName, ContrCod) <> 0 Then GoTo Found    '  ���� � �������?
'''                    If IsRightContrOppLink(OppId, ContrCod) Then GoTo Found
'''                End If
'''            Next i
'''    ' -- ���� ��������, ���� ��� ������� SF �����������, �� ������� �������� �� �����
'''    '    � ���� ������ �������� ��������� ������ ���������� ��� ����� � ���������
'''        End If
'''
'''        For i = 2 To EOL_SFopp
'''            If .Cells(i, 4) = Account Then
'''                OppN = .Cells(i, SFOPP_OPPN_COL)
'''                OppT = .Cells(i, SFOPP_TYP_COL)
'''                OppCur = .Cells(i, SFOPP_TO_PAY_CUR_COL)
'''                OppToPayRub = .Cells(i, SFOPP_TO_PAY_VAL_COL) * CurRate(OppCur)
'''                OppCloseDate = .Cells(i, SFOPP_CLOSEDATE_COL)
'''                OppId = .Cells(i, SFOPP_OPPID_COL)
'''                If InStr(OppT, SeekOppType) <> 0 _
'''                        And IsSameTeam(Sale, .Cells(i, SFOPP_SALE_COL), OppN) _
'''                        And OppToPayRub >= Rub _
'''                        And Dat <= OppCloseDate Then
'''                    If .Cells(i, SFOPP_PROBABILITY_COL) <> 0 Then
'''                        GoTo Found
'''                    Else
'''                        Msg = "� ����������� '" & Account & "'" _
'''                            & vbCrLf & vbCrLf & "���� ������ Closed/Lost" _
'''                            & vbCrLf & vbCrLf & OppName _
'''                            & vbCrLf & vbCrLf & "���������� ��� ���?"
'''                        Respond = MsgBox(Msg, vbYesNoCancel)
'''                        If Respond = vbCancel Then ExRespond = False
'''                        If Respond = vbYes Then
'''                            ErrMsg WARNING, "!! ��������� ��������� ������� " & OppN _
'''                                & vbCrLf & vbCrLf & "� ���� ������� ������!"
'''                            GoTo Found
'''                        End If
'''                    End If
'''                End If
'''            End If
'''        Next i
'''    End With
'''    Exit Function
'''
'''Found:
'''    IsOpp = DB_SFDC.Sheets(SFopp).Cells(i, SFOPP_OPPID_COL)    ' ���� ������ ������� ����
'''
'''    Const ErMsg = " <!> WARNING IsOpp: ���������� ������� ������� Id= "
'''    If OppToPayRub < Rub Then _
'''        LogWr ErMsg & OppId & "(" & OppN & ") ������ " & Rub & " ��������� ������ �������"
'''    If OppCloseDate < Dat Then _
'''        LogWr ErMsg & OppId & "(" & OppN & ") ���� " _
'''            & DDMMYYYY(Dat) & " ����� ���� �������� ������� " & DDMMYYYY(OppCloseDate)
'''    If InStr(OppT, SeekOppType) = 0 Then _
'''        LogWr ErMsg & OppId & "(" & OppN & ") ��� ������� '" & T _
'''            & "' �� ������������� ���� ������� '" & OppT & "'"
'''End Function
'''Sub ContrOppLink(iPay, ContrK, ContrId, OppId)
''''
'''' �������� ����� ������� - ������ �� ������� � ������ iPay
''''   25.3.12
''''   25.4.12 - replace Vlookup �� ����
'''
'''    Dim Contr_Opp, Contr_Pay
'''    Dim PayCode, OpNm, Resp, Msg As String
'''
'''    If ContrK = "" Then Exit Sub
'''    PayCode = Sheets(PAY_SHEET).Cells(iPay, PAYCODE_COL)
'''    Contr_Opp = ContrOppN(ContrId)
'''    Contr_Pay = OppNbyId(OppId)
'''    If Contr_Opp = Contr_Pay Then Exit Sub
'''    OpNm = ""
'''    Dim i As Integer
'''    For i = 2 To EOL_SFopp
'''        If OppId = Sheets(SFopp).Cells(i, SFOPP_OPPID_COL) Then
'''            OpNm = Sheets(SFopp).Cells(i, SFOPP_OPPNAME_COL)
'''        End If
'''    Next i
'''    Msg = "������� '" & ContrK & "'"
'''    If Contr_Opp = 0 Then
'''        Msg = Msg & " �� ������ � ��������."
'''    Else
'''        Msg = Msg & "������ � �������� '" _
'''            & OppNameByN(Contr_Opp) & "' (" & Contr_Opp & ")"
'''    End If
'''    Msg = Msg & vbCrLf & vbCrLf _
'''        & "������������ ����� '" _
'''        & OpNm & "' (" & Contr_Pay & ")"
'''    If Contr_Opp = 0 Then
'''        Msg = Msg & vbCrLf & vbCrLf & "������ ?"
'''    Else
'''        Msg = Msg & vbCrLf & vbCrLf & "������� ?"
'''    End If
'''    Resp = MsgBox(Msg, vbYesNoCancel)
'''    If Resp = vbCancel Then ExRespond = False
'''    If Resp = vbYes Then
'''        EOL_ContrLnk = EOL_ContrLnk + 1
'''        With Sheets(NewContractLnk)
'''            .Cells(EOL_ContrLnk, 1) = ContrId
'''            .Cells(EOL_ContrLnk, 2) = OppId
'''        End With
'''    End If
'''End Sub
'''
'''Sub P_PaidContract()
''''
'''' ��������� ����� �������� �� ���������, ������������ ����� ����� ��� Payment
''''   5.1.2012
''''   10.1.2012 - ������������ P_PaidContract �� ����� ��������
''''   14.2.2012 - ������ ������ ����� � �������� �� ��������
'''
'''    Const WSheetName = "P_PaidContract" ' ��� �������� ����� � Excel
'''
'''    ModStart WSheetName, "����� ������� �� ��������� � ����� " & WSheetName
'''
'''' ������� ���� ��������
'''    AutoFilterReset 1
'''
'''    Set Payments = ActiveSheet.Range("A1:AC" & ActiveSheet.UsedRange.Rows.Count)
'''    Payments.AutoFilter Field:=25, Criteria1:="<>"  ' ������� 1�, ��� ������ �������
'''    Payments.AutoFilter Field:=4, Criteria1:="="    '   .. ������� ��� ��� � SF
'''    Payments.AutoFilter Field:=1, Criteria1:="<>"   '       .. � ����������� ����
'''
'''    Range("B1:B" & ActiveSheet.UsedRange.Rows.Count - 3).Copy ' �������� ��������� �������
'''
'''    Sheets(WSheetName).Activate
'''    Range("A3").Select
'''    ActiveSheet.Paste       ' �������� ��������� ������� ����� 2� ������ P_PaidContract
'''
'''    Lines = Sheets(WSheetName).UsedRange.Rows.Count - 6
'''    Range("C2:P" & Lines + 2).Select
'''    Selection.FillDown                  ' ��������� ������� ������
'''    Selection.RowHeight = 15
'''
'''    Rows("2:3").Delete                  ' ������� ������ ������ ������ � ��������� �� ��������
'''    Range("A2:A" & Lines).Interior.Color = Gray
'''    Rows(Lines + 1 & ":1000").Delete
'''
'''' ????? ������-�� �� �������� ������ - �������� � �������� OppId � Excel - �������� �������
''''    Set PaidContr = ActiveSheet.Range("A1:I" & ActiveSheet.UsedRange.Rows.Count)
''''    PaidContr.AutoFilter Field:=9, Criteria1:="<>0"       ' ������� 1�, ��� ������ �������
'''
'''    For i = 2 To Lines
'''        If Range("I" & i).value = 0 Then Rows(i).Hidden = True
'''    Next i
'''
'''    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Payment"
'''    WritePaid "Paid.txt"
'''    Shell "quota2.bat"
'''
'''    AutoFilterReset 1
'''    ModEnd WSheetName
'''End Sub

Sub testGoodType()
    Dim res(1 To 5) As String, Flg(4) As Boolean
    Set DB_MATCH = FileOpen(F_MATCH)
    
    Call GoodType("AutoCAD Subscription Renewal/3", Flg)
    
    res(1) = GoodType("xxx ������� xxx")    ' ������������  (������� '�������[ $]'
                                            ', �.�. ����� �����, ��� �������)
    res(2) = GoodType("xx ���� ")           ' �� ������������, �.�. ������
    res(3) = GoodType("xxx ��������")       ' ����������    (������� '�������')
                                            '   �.�. ����� ����� - �� ����� 1 ������������� �������
    res(4) = GoodType("xx��x �������")
    res(5) = GoodType("3D ����������� SpacePilot PRO,�black,�USB,�CAD�Professional/1;")
    If res(1) <> "������������" Then Stop
    If res(2) <> "� � � � � �" Then Stop
    If res(3) <> "����������" Then Stop
    If res(4) <> "������" Then Stop
    
    Stop
 
End Sub
Function GoodType(ByVal G As String) As String
'
' - GoodType(Good, [IsLicense, IsSubscription, IsRenewal, IsWork)
'       ���������� ��� ������ Goods �� ������� � We
'       � ����� ��������, ��������, Renewal  � ������.
'       ���� ���������� ��� �� ������ - ������ � GoodType = ""
'   19.2.2012
'   5.10.12 - ������ ������ � DB_MATCH �� We
'   18.12.12 - LCase(G)
'   20.12.12 - ���������� ������ ����� � ������� Goods
'   22.12.12 - ��������� � InStr �������� �� ��������� � patTest
'               (RegExp, ���������� ���������) �.����
'   25.12.12 - patTest ���������� ������ ���� � ������ �������� �������� $, ����� InStr

    Dim j As Integer
    Dim iG As Range
    Dim S, Goods() As String, GoodW As String
    
    GoodType = ""
    If G = "" Then Exit Function
    G = LCase(Trim(G))
    For Each iG In DB_MATCH.Sheets(We).Range("Goods").Rows
        GoodType = iG.Cells(1, 1)
        S = LCase(iG.Cells(1, 2))
        Goods = Split(S, ",")   ' � Goods ������ ������� ������� ����
        For j = 0 To UBound(Goods)
            GoodW = Trim(Goods(j))
            If GoodW <> "" Then
                If Left(GoodW, 1) = "$" Then
                    If patTest(G, Mid(GoodW, 2)) Then GoTo Found
                Else
                    If InStr(G, GoodW) > 0 Then GoTo Found
                End If
            End If
        Next j
    Next iG
    ErrMsg FATAL_ERR, "����������� ��� ������ " & G
''''    End
Found:
End Function
    Sub testGoodJob()
    Set DB_MATCH = FileOpen(F_MATCH)
    Dim i As Long
    t0 = 0
    t1 = 0
    t2 = 0
    
    For i = 1 To 1000
    Call GoodJob("��������� xxx Zprinter xxx", "������������", 4)
    Next i
End Sub
Function GoodJob(Good As String, GoodType As String, JobType As Long) As Boolean

'
' - GoodJob(Good,GoodType,JobType)  - ���������� True ���� ����� Good ���� GoodType
'                                     ������������� ���� ����� JobType
' 7.1.2013
' 3.2.13 ��������������

    Const JOB_COL0 = 12
    
    Dim iGoodType As Long, Rng As Range, i As Long
    Dim Goods() As String, GoodW As String
    Dim Pattern As String
    Dim t1a As Single
    
    GoodJob = False
       
    If Good = "" Then Exit Function
t1a = Now
    With DB_MATCH.Sheets(We).Range("Goods")
        Set Rng = .Columns(1)
'''        iGoodType = Application.Match(GoodType, Rng, 0) '����� ������ �� ���� ������
        For iGoodType = 1 To Rng.Rows.Count
            If GoodType = .Cells(iGoodType, 1) Then GoTo FoundType
        Next iGoodType
        ErrMsg FATAL_ERR, "�������� ��� ������ '" & GoodType & "'"
FoundType:
        Pattern = .Cells(iGoodType, CLng(JobType) + JOB_COL0)
        Goods = Split(Pattern, ",")
    End With
t1 = t1 + (Now - t1a)
    For i = 0 To UBound(Goods)
        GoodW = Trim(Goods(i))
        If GoodW <> "" Then
            If Left(GoodW, 1) = "$" Then
                If patTest(Good, Mid(GoodW, 2)) Then GoTo Found
            Else
                If InStr(Good, GoodW) > 0 Then GoTo Found
            End If
        End If
    Next i
t2 = t2 + Now - t1a
    Exit Function
Found:
    GoodJob = True
End Function

Function IsSubscription(Good, GT) As Boolean
'
' ���������� True, ���� ����� - ��������/Subscription/Maintanence
' � ����������� �� ���� ������ GT. ����� - �������� ��������, �.�. False.
' ������������ �� ������� ������� � We
'   24.4.2012

    Dim Sbs As String
    Dim iG As Range

    Const SBSCOL = 7
    
    IsSubscription = False
    
    Sbs = ""
    For Each iG In DB_MATCH.Sheets(We).Range("GoodSbs").Rows
        If iG.Cells(1, 1) = GT Then
            Sbs = iG.Cells(1, SBSCOL)
            Exit For
        End If
    Next iG
    If Sbs = "" Then Exit Function
    
    If Sbs = "TRUE" Then
        IsSubscription = True
        Exit Function
    End If
    
    Dim i As Integer
    Dim SbsWords() As String
    Dim LGood As String
    LGood = LCase$(Good)
    
    SbsWords = Split(LCase$(Sbs), ",")
    For i = LBound(SbsWords) To UBound(SbsWords)
        If InStr(LGood, Trim(SbsWords(i))) > 0 Then
            IsSubscription = True
            Exit Function
        End If
    Next i
End Function
Function IsWork(ByVal Good As String) As Boolean
'
' - IsWork(Good)    - ���������� True, ���� ����� - ������
' 29.10.12

    Dim Wrd() As String, Wokabulary As String, i As Long
    
'    For Each iG In Range("WorksTable").Rows
'        Ent = Split(LCase$(Good), ",")
'
'    Next iG
'
    Good = LCase(Good)
    Wokabulary = DB_MATCH.Sheets(We).Range("WorksTable").Cells(1, 2)
    Wrd = Split(LCase(Wokabulary), ",")
    IsWork = True
    For i = LBound(Wrd) To UBound(Wrd)
        If InStr(Good, Trim(Wrd(i))) > 0 Then Exit Function
    Next i
    IsWork = False
End Function
Function TypOpp(Good) As String
'
' - TypOpp(Good) - ������������� ���� ������� �� ������������ ������
'
' 29.10.12
'  9.9.13 - ������� ��������� - GoodType ���������� �� ������������
    
    Dim WeRange As Range, i As Long, iG As Range, GoodTp As String
    Set WeRange = DB_MATCH.Sheets(We).Range("GoodSbs")

    TypOpp = ""
    GoodTp = GoodType(Good)
    For Each iG In WeRange.Rows
        If iG.Cells(1, 1) = GoodTp Then
            TypOpp = iG.Cells(1, 11)
            If TypOpp <> "" Then Exit Function
            If IsWork(Good) Then
                TypOpp = "������"
                Exit Function
            End If
            If IsSubscription(Good, GoodTp) Then
                TypOpp = "��������"
            Else
                TypOpp = "��������"
            End If
            Exit For
        End If
    Next iG
End Function
