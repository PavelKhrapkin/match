Attribute VB_Name = "PaidAnalitics"
'---------------------------------------------------------------------------
' PaidAnalitics -- ������� ��� ������� ���������� ��������
'
' * PaidHandling()  - ������ �� ����� �������� 1�, ��������� � SF
' - GoodType(Good)              - ���������� ������ - ��� ������ Good
' - IsSubscription(Good, GT)    - ���������� True, ���� ����� - ��������
'
'   30.5.2012

Option Explicit

Const MinNewOpp = 120000
Const MinNewOppDialog = 200000  ' ���� ��� ��������� ����������� � �����
Sub NewPaidDog()
'
' S
'
    StepIn
    
    Dim P As TOCmatch
    Dim i As Long
    
    P = GetRep(PAY_SHEET)
    With DB_1C.Sheets(PAY_SHEET)
        For i = 2 To P.EOL
            Progress i / P.EOL
            If .Cells(i, PAYISACC_COL) <> "" And .Cells(i, PAYINSF_COL) = "" Then
                If .Cells(i, PAYDOGOVOR_COL) <> "" Then
                    WrNewSheet NEW_PAYMENT, DB_1C.Sheets(PAY_SHEET), i
                End If
            End If
        Next i
    End With
End Sub
Sub PaidHandling()
'
' ������ �� �������� � ����� ����� �������� �� ���������� ��������
'   16.8.12 match 2.0

'---- ���������� ��������� ���������� ----------
' � ������������ ���� "����' - �������� ���� �� ����� ��������1�
'                        =   - ��������� ���������� � ���� ���������
    Dim PayK As String      '���� - ��� �������
    Dim Dat As Date         '���� - "���� ����.�����"
    Dim Acc As String       '���� - "������"
    Dim Rub As Variant      '���� - "���� ���"
    Dim Sale As String      '���� - "��������"
    Dim good As String      '���� - "������" ����� ��������, ������������
    Dim t As String         ' = ��� ������ �� ������������
    Dim Sbs As Boolean      ' = True ���� ������������ �������� ��������
    Dim Dogovor As String   '���� - "�������"
    Dim MainDog As String   '���� - "���.�������"
    Dim ContrK As String    ' = ����� <���.�������>/<�������>
    Dim ContrId As String   ' = Id �������� ��� � SF
    Dim OppId As String     ' = Id ������� � SF
    
    Dim i                   ' = ������ ��������� ����� �� ��������
    Dim t0, Tbeg, TI    ' ��� �������������� �� �������
    t0 = Timer
'-----------------------------------------------

    Dim ts1 As Long, ts2 As Long, ts3 As Long, ts4 As Long, ts5 As Long ' profiling ?????????
    Dim ts1S As Long, ts2S As Long, ts3S As Long, ts4S As Long, ts5S As Long ' profiling ?????????
    ts1 = 0: ts2 = 0: ts3 = 0: ts4 = 0: ts2 = 0: ts5 = 0

    TI = Timer
    LogWr t0 - TI & " PaidAnalitics: ������"
    Dim SumNewPay
    SumNewPay = 0
    
    
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim P As TOCmatch, Opp As TOCmatch
    GetRep TOC
    DB_MATCH.Sheets(We).Activate
    ClearSheet O_NewOpp, Range("HDR_NewOpp")
    ClearSheet P_Paid, Range("HDR_Payment")
    Opp = GetRep(SFopp)
    EOL_SFopp = Opp.EOL
    P = GetRep(PAY_SHEET)
    
    For i = 2 To P.EOL
        Progress (i / P.EOL)
        If ExRespond = False Then GoTo Ex
        With DB_1C.Sheets(P.SheetN)
'''''            .Activate
        ' Account � SF ����, ������� � SF ���, ��� � �������� �� ������
            Acc = Compressor(.Cells(i, PAYACC_COL)) ' �����������
            If .Cells(i, PAYISACC_COL) <> "" And _
                    Trim(.Cells(i, PAYDOC_COL)) <> "" And _
                    Trim(.Cells(i, PAYSALE_COL)) <> "" Then
                Dat = .Cells(i, PAYDATE_COL)    ' ���� �������
                Sale = .Cells(i, PAYSALE_COL)   ' ��������
                good = .Cells(i, PAYGOOD_COL)   ' �����
                t = GoodType(good)              ' ��� ������ �� �����
                Sbs = IsSubscription(good, t)   ' ��������?
                Rub = .Cells(i, PAYRUB_COL)     ' ����� ������� ���
                Dogovor = .Cells(i, PAYDOGOVOR_COL)
                MainDog = Mid(.Cells(i, PAYOSNDOGOVOR_COL), 9)
                ContrK = ContrCod(Dogovor, MainDog)
                ContrId = ContractId(ContrK)    ' Id ��������, ���� ����
                Dim IsInSF As String
                IsInSF = ""

    '================ ������ ��� �� ������� � SF? =============================
                If .Cells(i, PAYINSF_COL) = "" Then
                    
                    ts1 = Timer                     ' tttttttttttttttttttttttttttttttttttttt
                    
                    OppId = IsOpp(Sale, Acc, t, Rub, Dat, ContrK) ' Id ������� � SF
                    If OppId = "" Then
                        NewOpp Acc, ContrK, Dat, Sale, Rub, "RUB", t, Sbs
                    Else
            '>>>>  ��������� ������ �������
                        NewPay i, OppNbyId(OppId), ContrId
            '>>>>  ����� ������� � ��������� ��� �������� ������ �������� � SF
                        If ContrK <> "" And ContrId = "" Then
                            NewContract Dogovor, MainDog, ContrK
                        Else
                            ContrOppLink i, ContrK, ContrId, OppId
                        End If
                    End If
                    ts1S = ts1S + (Timer - ts1)      ' tttttttttttttttttttttttttttttttttttt
                    ts1 = ts1
                End If
            End If
        End With
    Next i
    
'    MsgBox "SumNewPay = " & SumNewPay
    
    Dim ResultMsg
Ex: ResultMsg = "����: �������� " & EOL_NewPay - 1 & " ����� ��������; " _
        & EOL_NewOpp - 1 & " ����� ��������; " _
        & EOL_PaymentUpd - 1 & " �������� ������� � ����������; " _
        & EOL_ContrLnk - 1 & " ��������� ������� � ���������; " _
        & EOL_AdAcc - 1 & " ����� �����������;" _
        & EOL_ADSKlnkPay - 1 & " �������� ������� � ����������� Autodesk;"
    LogWr ResultMsg
    MsgBox ResultMsg
 
    MsgBox "time1: " & ts1S & "; time2: " & ts2S & "; time3: " & ts3S & "; time4: " & ts4S
      
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Payment\"
    WriteCSV P_Paid, "Paid.txt"
    Shell "quota2.bat"
    WriteCSV P_PaymentUpd, "PmntUpd.txt"
    Shell "quota3.bat"
    
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\OppInsert\"
    WriteCSV O_NewOpp, "OppInsert.txt"
    Shell "quota2.bat"
    
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Dogovor\"
    WriteCSV NewContractLnk, "ContrUpd.txt"
    Shell "quota3.bat ContrUpd.txt C:\SFconstr\Dogovor\ContrUpd.csv"
    
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Account\"
    WriteCSV A_Acc, "AdAcc.txt"
    Shell "quota_Acc.bat"

    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\ADSK\P_ADSK"
    WriteCSV P_ADSKlink, "P_ADSKlink.txt"
    Shell "quota_P_ADSK.bat"

'''    ModEnd 1
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
Function IsOpp(Sale, Account, t, Rub, Dat, ContrCod)
'
' ��������, ���� �� � ����������� Account �� ��������� ���������� ������ ���� �.
' ���� ������ �� �������� �� �����, ������� Rub - ���������� Id ����� �������.
' ��� �� ������������ ������� ���������� ������ ��� ����� ��� � ���������,
' ��������� � TargetContrK �� ������������ ������ ����� �������� � �������� �������.
' ���� ��� - ���������� ""
'   13.2.2012
'   21.2.2012 - ������� �� �������� ��������
'   28.12.12 - ����� ����������� ������� ��� �������� �������
'   3.3.12 - ��������� ���� �������� ������� OppCrDate,
'   5.3.12 - Optional ContrCod ��� �������� ������� ����� ������ ������� � SF
'   9.3.12 - �� ���� ������, ���� ������� �������� �� �������� ����� �
'  16.3.12 - ������� ������� Closed/Lost, �� ���� ����������� = 0%
'  18.3.12 - ���������� Id � �� ����� ���������� �������
'  19.3.12 - TargetContrK - ������� ������� ��� ����� � �������� Opp
'  22.3.12 - ������� ��� � � ����� ������� - ���� ������ - ��� ������������ Opp
'  31.3.12 - ���������� ������ ������ ����������� �������
'  20.4.12 - bug fix ��� �������� � ���������
'  25.4.12 - ������ � ��������� �������������� � ��������� ������� �� ��������
'  30.5.12 - ��������� � ��������� ������ "������������ �������� � �������"
'  31.5.12 - ��������� �������� ������ � Close Lost ������

    Dim i, OppToPayRub, OppCur, OppN, ContrOpp, OppName, OppId
    Dim OppCloseDate As Date, OppCrDat As Date
    Dim Msg As String           ' ��������� �� ������ � IsOpp
    Dim Respond As String       ' ����� � ������������� �������
    
    IsOpp = ""
    
    If t = "" Then Exit Function
    If Not IsRightSale(Sale, t) Then
        ErrMsg FATAL_ERR, "������������ �������� " & Sale & " � ������� " & t
        Exit Function
    End If
    
    Dim SeekOppType As String, OppT As String
    SeekOppType = WorksheetFunction.VLookup(t, Range("OppTypeRng"), 4, False)
 
    With DB_SFDC.Sheets(SFopp)
'-- ���� ������ ������ � ��������� -- ��������, ��� ����� � ContrCod ����
        If ContrCod <> "" Then
            For i = 2 To EOL_SFopp
                If .Cells(i, 4) = Account Then
                    OppName = .Cells(i, SFOPP_OPPNAME_COL)          ' ��� ��������
                    OppId = .Cells(i, SFOPP_OPPID_COL)
                    If InStr(OppName, ContrCod) <> 0 Then GoTo Found    '  ���� � �������?
                    If IsRightContrOppLink(OppId, ContrCod) Then GoTo Found
                End If
            Next i
    ' -- ���� ��������, ���� ��� ������� SF �����������, �� ������� �������� �� �����
    '    � ���� ������ �������� ��������� ������ ���������� ��� ����� � ���������
        End If

        For i = 2 To EOL_SFopp
            If .Cells(i, 4) = Account Then
                OppN = .Cells(i, SFOPP_OPPN_COL)
                OppT = .Cells(i, SFOPP_TYP_COL)
                OppCur = .Cells(i, SFOPP_TO_PAY_CUR_COL)
                OppToPayRub = .Cells(i, SFOPP_TO_PAY_VAL_COL) * CurRate(OppCur)
                OppCloseDate = .Cells(i, SFOPP_CLOSEDATE_COL)
                OppId = .Cells(i, SFOPP_OPPID_COL)
                If InStr(OppT, SeekOppType) <> 0 _
                        And IsSameTeam(Sale, .Cells(i, SFOPP_SALE_COL), OppN) _
                        And OppToPayRub >= Rub _
                        And Dat <= OppCloseDate Then
                    If .Cells(i, SFOPP_PROBABILITY_COL) <> 0 Then
                        GoTo Found
                    Else
                        Msg = "� ����������� '" & Account & "'" _
                            & vbCrLf & vbCrLf & "���� ������ Closed/Lost" _
                            & vbCrLf & vbCrLf & OppName _
                            & vbCrLf & vbCrLf & "���������� ��� ���?"
                        Respond = MsgBox(Msg, vbYesNoCancel)
                        If Respond = vbCancel Then ExRespond = False
                        If Respond = vbYes Then
                            ErrMsg WARNING, "!! ��������� ��������� ������� " & OppN _
                                & vbCrLf & vbCrLf & "� ���� ������� ������!"
                            GoTo Found
                        End If
                    End If
                End If
            End If
        Next i
    End With
    Exit Function

Found:
    IsOpp = DB_SFDC.Sheets(SFopp).Cells(i, SFOPP_OPPID_COL)    ' ���� ������ ������� ����
        
    Const ErMsg = " <!> WARNING IsOpp: ���������� ������� ������� Id= "
    If OppToPayRub < Rub Then _
        LogWr ErMsg & OppId & "(" & OppN & ") ������ " & Rub & " ��������� ������ �������"
    If OppCloseDate < Dat Then _
        LogWr ErMsg & OppId & "(" & OppN & ") ���� " _
            & DDMMYYYY(Dat) & " ����� ���� �������� ������� " & DDMMYYYY(OppCloseDate)
    If InStr(OppT, SeekOppType) = 0 Then _
        LogWr ErMsg & OppId & "(" & OppN & ") ��� ������� '" & t _
            & "' �� ������������� ���� ������� '" & OppT & "'"
End Function
Sub ContrOppLink(iPay, ContrK, ContrId, OppId)
'
' �������� ����� ������� - ������ �� ������� � ������ iPay
'   25.3.12
'   25.4.12 - replace Vlookup �� ����

    Dim Contr_Opp, Contr_Pay
    Dim PayCode, OpNm, Resp, Msg As String
    
    If ContrK = "" Then Exit Sub
    PayCode = Sheets(PAY_SHEET).Cells(iPay, PAYCODE_COL)
    Contr_Opp = ContrOppN(ContrId)
    Contr_Pay = OppNbyId(OppId)
    If Contr_Opp = Contr_Pay Then Exit Sub
    OpNm = ""
    Dim i As Integer
    For i = 2 To EOL_SFopp
        If OppId = Sheets(SFopp).Cells(i, SFOPP_OPPID_COL) Then
            OpNm = Sheets(SFopp).Cells(i, SFOPP_OPPNAME_COL)
        End If
    Next i
    Msg = "������� '" & ContrK & "'"
    If Contr_Opp = 0 Then
        Msg = Msg & " �� ������ � ��������."
    Else
        Msg = Msg & "������ � �������� '" _
            & OppNameByN(Contr_Opp) & "' (" & Contr_Opp & ")"
    End If
    Msg = Msg & vbCrLf & vbCrLf _
        & "������������ ����� '" _
        & OpNm & "' (" & Contr_Pay & ")"
    If Contr_Opp = 0 Then
        Msg = Msg & vbCrLf & vbCrLf & "������ ?"
    Else
        Msg = Msg & vbCrLf & vbCrLf & "������� ?"
    End If
    Resp = MsgBox(Msg, vbYesNoCancel)
    If Resp = vbCancel Then ExRespond = False
    If Resp = vbYes Then
        EOL_ContrLnk = EOL_ContrLnk + 1
        With Sheets(NewContractLnk)
            .Cells(EOL_ContrLnk, 1) = ContrId
            .Cells(EOL_ContrLnk, 2) = OppId
        End With
    End If
End Sub

Sub P_PaidContract()
'
' ��������� ����� �������� �� ���������, ������������ ����� ����� ��� Payment
'   5.1.2012
'   10.1.2012 - ������������ P_PaidContract �� ����� ��������
'   14.2.2012 - ������ ������ ����� � �������� �� ��������
    
    Const WSheetName = "P_PaidContract" ' ��� �������� ����� � Excel
    
    ModStart WSheetName, "����� ������� �� ��������� � ����� " & WSheetName
    
' ������� ���� ��������
    AutoFilterReset 1
    
    Set Payments = ActiveSheet.Range("A1:AC" & ActiveSheet.UsedRange.Rows.Count)
    Payments.AutoFilter Field:=25, Criteria1:="<>"  ' ������� 1�, ��� ������ �������
    Payments.AutoFilter Field:=4, Criteria1:="="    '   .. ������� ��� ��� � SF
    Payments.AutoFilter Field:=1, Criteria1:="<>"   '       .. � ����������� ����
       
    Range("B1:B" & ActiveSheet.UsedRange.Rows.Count - 3).Copy ' �������� ��������� �������
    
    Sheets(WSheetName).Activate
    Range("A3").Select
    ActiveSheet.Paste       ' �������� ��������� ������� ����� 2� ������ P_PaidContract
    
    Lines = Sheets(WSheetName).UsedRange.Rows.Count - 6
    Range("C2:P" & Lines + 2).Select
    Selection.FillDown                  ' ��������� ������� ������
    Selection.RowHeight = 15
    
    Rows("2:3").Delete                  ' ������� ������ ������ ������ � ��������� �� ��������
    Range("A2:A" & Lines).Interior.Color = Gray
    Rows(Lines + 1 & ":1000").Delete
    
' ????? ������-�� �� �������� ������ - �������� � �������� OppId � Excel - �������� �������
'    Set PaidContr = ActiveSheet.Range("A1:I" & ActiveSheet.UsedRange.Rows.Count)
'    PaidContr.AutoFilter Field:=9, Criteria1:="<>0"       ' ������� 1�, ��� ������ �������

    For i = 2 To Lines
        If Range("I" & i).Value = 0 Then Rows(i).Hidden = True
    Next i
    
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Payment"
    WritePaid "Paid.txt"
    Shell "quota2.bat"
    
    AutoFilterReset 1
    ModEnd WSheetName
End Sub
Function GoodType(G) As String
'
' ���������� ��� ������ G �� ������� � We.
' ���� ���������� ��� �� ������ - ������ � GoodType = ""
'   19.2.2012

    Dim j As Integer
    Dim iG As Range
    Dim S, Goods() As String
    
    GoodType = ""
    If G = "" Then Exit Function
    For Each iG In Range("Goods").Rows
        GoodType = iG.Cells(1, 1)
        S = iG.Cells(1, 2)
        Goods = split(S, ",")   ' � Goods ������ ������� ������� ����
'If GoodType = "� � � � � �" Then
'j = j
'End If
        For j = 0 To UBound(Goods)
            If InStr(G, Trim(Goods(j))) > 0 Then Exit Function
        Next j
    Next iG
    ErrMsg FATAL_ERR, "����������� ��� ������ " & G
End Function
Function IsSubscription(good, GT) As Boolean
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
    For Each iG In Range("GoodSbs").Rows
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
    LGood = LCase$(good)
    
    SbsWords = split(LCase$(Sbs), ",")
    For i = LBound(SbsWords) To UBound(SbsWords)
        If InStr(LGood, Trim(SbsWords(i))) > 0 Then
            IsSubscription = True
            Exit Function
        End If
    Next i
    
End Function
