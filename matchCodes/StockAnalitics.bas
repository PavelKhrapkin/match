Attribute VB_Name = "StockAnalitics"
'---------------------------------------------------------------------------------
' StockAnalitics  - ������ ���� �� ������
'
' [*] StockHandling()   - ������ �� ����� "�����", ����� ��������� Autodesk
'  -  FindAcc1C(Client, Acc1C) - ����� ����� 1� ���������� � Client ��� Acc1C
'  -  SeekInv(Str) - ��������� ����� � ��������� ������ Str
'  -  SNhandl(Acc1C, PayN, StockSN) - ��������� SN � ��������� �����
'  ?  RemIgnoredSN(S)   - �������� ������������ ���� � ������ � S �� SN
' (*) Sndedup() - ������������ SN ��������� �� ������ - ��������!
'  -  SN_ADSKbyStock(PayId, Acc, Dat) - ���������� SN �������� ADSK �� ������

'   15.09.2013

Option Explicit

Sub StockHandling()
'
' [*] StockHandling()   - - ������ �� ����� "�����", ����� ��������� Autodesk
'   7.5.12
'  13.5.12 - � ������� ���� ������ ����� ���� �� ��������, � �� ����� ������
'  13.5.12 - ��������� ����������� ������� SN_SF - SN ��������� � SF
'  23.5.12 - �������
'   7.6.12 - ��� ��������� ����� ������� FindAcc � FindPayN ��� ���������� Client

    Dim i As Integer
    Dim PayN As Integer
    Dim PayId As String     ' = Id ������� � SF
    Dim Client As String    ' ������ � ��������� �����, ������, ����, �����
    Dim SameClient As Boolean
    Dim Acc1C As String     ' ��� ����������� � ����������� 1�
    Dim Dat As Date         '���� "����" � ��������� �����
    Dim Good, t As String   ' ����� (������������) � ��� ������
    Dim StockSN As String   ' ��������� ������ �� SN
    Dim SNinSF As String    ' SN ��� ���������� � SF
    Dim NewSN As String     ' SN, �������� ��� ��� � SF
    Dim ContrADSK As String ' �������� ADSK - �� SF ��� �� ������
    
    Lines = ModStart(STOCK_SHEET, "������ �� ������: SN Autodesk", True)
    CheckSheet STOCK_SHEET, 1, STOCK_PRODUCT_COL, STOCK_STAMP

    With Sheets(STOCK_SHEET)
        For i = 2 To Lines
            Progress i / Lines
            If ExRespond = False Then Exit For
If i >= 7766 Then
i = i
End If
            Client = .Cells(i, STOCK_CLIENT_COL)
            Dat = TxDate(.Cells(i, STOCK_DATE_COL))
            If Client = .Cells(i - 1, STOCK_CLIENT_COL) Then
                SameClient = True
            Else
                SameClient = False
                Acc1C = FindAcc(Client)                 ' ����� �����������
                PayN = FindPayN(Client, Acc1C, Dat)     ' ����� �����/�������
            End If
            .Cells(i, STOCK_ACC_COL) = Acc1C
            If PayN > 0 And PayN <= EOL_PaySheet Then
                .Cells(i, STOCK_INVOICE_COL) = _
                        Sheets(PAY_SHEET).Cells(PayN, PAYINVOICE_COL)
                Good = Sheets(PAY_SHEET).Cells(PayN, PAYGOOD_COL)
                t = GoodType(Good)              ' ��� ������ �� �����
                .Cells(i, STOCK_GOOD_COL) = t
                If t = WE_GOODS_ADSK Then
                    StockSN = Sheets(STOCK_SHEET).Cells(i, STOCK_SN_COL)
                    NewSN = SNhandl(Acc1C, PayN, StockSN, SNinSF, ContrADSK)
                    .Cells(i, STOCK_SF_SN_COL) = SNinSF
                    If SNinSF = "" Then
                        ContrADSK = GetContrADSKfrStock(StockSN)
                    Else
                        If ContrADSK <> "" Then
                            PayId = PayIdByK(Sheets(PAY_SHEET).Cells(PayN, PAYCODE_COL))
                            If IsADSK_PA(ContrADSK, PayId) Then
                                .Cells(i, STOCK_CONTRADSK_COL).Interior.Color = rgbLightGreen
                                .Cells(i, STOCK_INVOICE_COL).Interior.Color = rgbLightGreen
                            End If
                        End If
                    End If
                    Sheets(STOCK_SHEET).Cells(i, STOCK_CONTRADSK_COL) = ContrADSK
                    Sheets(STOCK_SHEET).Cells(i, STOCK_NEWSN_COL) = NewSN
                End If
            Else
                Sheets(STOCK_SHEET).Cells(i, STOCK_INVOICE_COL) = ""
            End If
        Next i
    End With
    
    ModEnd STOCK_SHEET
End Sub
Function FindPayN(Client, Acc1C, Dat) As Integer
'
' FindAcc1C(Client) - ����� ������� �� �����1� ����������� � Client ��� Acc1C
'   7.5.12
'  23.5.12 -�� ����� � Acc1C ����� ���� ��������� ����������� ����������� "+"
'           FindPayN ���� ���� �� ���� � ��������� � Acc1C ������ ��������
'  17.6.12 - ������� ����� ��- � ������ � ��������� ������������
'  15.9.13 ������������� PAYCANONAME_COL ������ PAYACC_COL

    Dim invoice As String
    Dim StockInv As String      '= ����, ���������� �� ���� ������ �� ������
    Dim D As Date               '���� "����" �� ������
    Dim AccInv As String        '= ��� �����������, ���������� �� �����
    Dim i1C As Integer          '= ����� ������ � ������� �� �����
    Dim Inv As String, Acc As String    '=
    Dim i As Integer, j As Integer      '=
    
    FindPayN = 0
    
    StockInv = SeekInv(Client)
    If StockInv = "" Then Exit Function
    If IsInv1C(StockInv, Dat, i1C) Then
        AccInv = Compressor(Sheets(PAY_SHEET).Cells(i1C, PAYCANONAME_COL))
        invoice = Sheets(PAY_SHEET).Cells(i1C, PAYINVOICE_COL)
        If Acc1C <> AccInv Then
            If Acc1C = "" Then
                Acc1C = AccInv
            Else
                ErrMsg WARNING, "�������� �� " & invoice & " " _
                    & AccInv & ", � �� ������ " & Acc1C
            End If
        End If
    End If
    
    If Acc1C = "" Or Acc1C = "*" Then Exit Function
    FindPayN = i1C
    
    
'----- ����� ����� � �������� -------
'
'    For i = 2 To EOL_PaySheet
'        Acc = Sheets(PAY_SHEET).Cells(i, PAYCANONAME_COL)
'        If InStr(Acc1C, Acc) <> 0 Then
'            If InStr(Invoice, StockInv) <> 0 Then
'                For j = 4 To 7 ' ��������� ������ ����� ���� "��-123 ..."
'                    If Mid(Invoice, j, 1) = " " Then
'                        Inv = Mid(Invoice, 4, j - 4)
'                        Exit For
'                    End If
'                Next j
'                Inv = "��-" & Inv
'                If Inv = StockInv Then
'                    FindPayN = i
'                    Acc1C = Acc
'                    Exit Function
'                End If
'            End If
'        End If
'    Next i
End Function
Function SeekInv(Str) As String
'
' - SeekInv(Str) - ��������� ����� � ��������� ������ Str
'   17.6.12

    Dim Sch As String
    Dim S As String
    Dim i As Integer

    Dim StWord() As String  '= ������ ���� � Str
    
    SeekInv = ""
    If Trim(Str) = "" Then Exit Function
    
    S = Replace(LCase(Str), "_", " ")
    S = Replace(LCase(S), ")", " ")
    S = Replace(LCase(S), "(", " ")
    S = Replace(LCase(S), """", " ")
    StWord = Split(S, " ")
    For i = LBound(StWord) To UBound(StWord)
        Sch = StWord(i)
        If Left(Sch, 1) = Chr(99) Or Left(Sch, 1) = "�" Then ' Ru ��� En "�"
            Do
                Sch = Mid(Sch, 2)
                If Sch = "" Then Exit Do
                If Left(Sch, 1) <> "-" And IsNumeric(Sch) Then GoTo FoundSeekInv
            Loop
            If i < UBound(StWord) Then
                Sch = StWord(i + 1) ' ������ ��-123 ����� ����� 'C 123'
                If IsNumeric(StWord(i + 1)) Then GoTo FoundSeekInv
            End If
        End If
    Next i
    Exit Function
FoundSeekInv:
    Sch = Abs(Sch)
    If Sch = 0 Or Sch >= 1000 Then Exit Function
    SeekInv = "��-" & Sch
End Function
Sub testSeekInv()
    Dim a(1 To 6) As String
    a(1) = SeekInv("����� ��� ""���-94"" ��-267 �� 07.10.11 ��������� ")
    a(2) = SeekInv("����� ��� ""���-94"" �-267 �� 07.10.11 ��������� ")
    a(3) = SeekInv("����� ��� ""���-94"" � -267 �� 07.10.11 ��������� ")
    a(4) = SeekInv("����� ��� ""���-94"" ��- 267 �� 07.10.11 ��������� ")
    a(5) = SeekInv("����� ��� ""���-94"" �� - 267 �� 07.10.11 ��������� ")  '!!! �� �����������!!!
End Sub
Function SeekPayN(ByVal Inv As String, ByVal Client As String, ByVal Dat As Date) As Long
'
' - SeekPayN(Inv, Client, Dat)  - ����������� ������ ������ � �������� �� ����� � ����
' 25.11.20

    Const INV_VALIDITY = 50                     'max ���� ����� ������ �����
    
    Dim Dic As TOCmatch, DicRange As Range      ' ������� �������� �����������
    Dim Acc As String, accWords() As String     ' ������ �������� � ����� � ����� �����������
    Dim Id As String, IdS() As String           ' ������ IdSFacc, ������������ + � ���������
    Dim IdSF(), N_IdSF As Long: N_IdSF = 0
''    Dim IdSF(1 To N_IdSF) As String             ' ������ IdSFacc �� ������ � �����
    
    Dim P As TOCmatch, Pdat As String, PayDat As Date
    Dim SFpaid As TOCmatch, SFRange As Range
    Dim InvSF As String
    Dim PayN As Long, i As Long, j As Long, N As Long
        
    SeekPayN = 0
    If Trim(Client) = "" Then Exit Function
                
    accWords = Split(RemIgnored(LCase$(Client)), " ")
    
    Dic = GetRep("DicAcc")
    
    With DB_TMP.Sheets(Dic.SheetN)
        Set DicRange = Range(.Cells(2, 1), .Cells(BIG, 3))
        For i = LBound(accWords) To UBound(accWords)
            Id = ""
            On Error Resume Next
            Id = WorksheetFunction.VLookup(accWords(i), DicRange, 3, False)
            On Error GoTo 0
            If Id = "" Then
                accWords(i) = ""
            Else
                IdS = Split(Id, "+")
                For j = LBound(IdS) To UBound(IdS)
                    ReDim Preserve IdSF(N_IdSF)
                    IdSF(N_IdSF) = IdS(j)
                    N_IdSF = N_IdSF + 1
                Next j
            End If
        Next i
    End With
    
    If N_IdSF = 0 Then Exit Function
    P = GetRep(PAY_SHEET)
    With DB_1C.Sheets(P.SheetN)
        Dim Prng As Range
        Set Prng = Range(.Cells(2, 1), .Cells(P.EOL, PAYINVOICE_COL))
        For i = 0 To N_IdSF - 1
            Do
                j = CSmatchSht(IdSF(i), 1, Prng, j)
                If j > 0 Then
                    If Left(.Cells(j, PAYINVOICE_COL), Len(Inv)) = Inv Then
'                    If InStr(.Cells(j, PAYINVOICE_COL), Inv) <> 0 Then
                        SeekPayN = j
                        Exit Function
                    End If
                End If
            Loop While j > 0
        Next i
''''''            Inv1C = ""
''''''            On Error Resume Next
''''''            Inv1C = WorksheetFunction.VLookup(IdSF(i), Prng, PAYINVOICE_COL, False)
''''''            On Error GoTo 0
''''''            If InStr(Inv1C, Inv) <> 0 Then
''''''                SeekPayN = WorksheetFunction.Match(IdSF(i), SFRange, 0)
''''''                SeekPayN = .Cells(1, SeekPayN)
''''''                Exit Function
''''''            End If
''''''        Next i
    End With
''i = i
''    SFpaid = GetRep(SF)
''    With DB_SFDC.Sheets(SF)
''        Set SFRange = Range(.Cells(2, 1), .Cells(BIG, SF_INV_COL))
''        For i = 1 To N_IdSF
''            InvSF = ""
''            On Error Resume Next
''            InvSF = WorksheetFunction.VLookup(IdSF(i), SFRange, SF_INV_COL, False)
''            On Error GoTo 0
''            If InStr(InvSF, Inv) <> 0 Then
''                SeekPayN = WorksheetFunction.Match(IdSF(i), SFRange, 0)
''                SeekPayN = .Cells(1, SeekPayN)
''                Exit Function
''            End If
''        Next i
''    End With
''i = i
'''
'''    P = GetRep(PAY_SHEET)
'''
'''    N = 1
'''    With DB_1C.Sheets(PAY_SHEET)
'''        Do
'''            PayN = 0
'''            On Error Resume Next
'''            PayN = Application.Match(Inv, _
'''                Range(.Cells(N, PAYINV_COL), .Cells(BIG, PAYINV_COL)), 0) _
'''                + N - 1
'''            Pdat = .Cells(PayN, PAYDATE_COL)
'''            On Error GoTo 0
'''            If IsEmpty(PayN) _
'''                    Or Not IsNumeric(PayN) _
'''                    Or PayN <= 0 _
'''                    Or Not IsDate(Pdat) Then Exit Function
'''            If Not IgnoredFirm(.Cells(PayN, PAYFIRM_COL)) Then
'''                Acc = LCase(.Cells(PayN, PAYCANONAME_COL))
'''                PaccW = split(RemIgnored(Acc), " ")
'''                For i = LBound(accWords) To UBound(accWords)
'''                    If accWords(i) <> "" Then
'''                        For j = LBound(PaccW) To UBound(PaccW)
'''                            If accWords(i) = PaccW(j) Then GoTo Found
'''                        Next j
'''                    End If
'''                Next i
'''                GoTo NextInv
'''Found:          PayDat = Pdat
'''                If Dat - PayDat < INV_VALIDITY And Dat >= PayDat Then
'''                    SeekPayN = PayN     ' ����� ����� ������ ������� PayN
'''                    Exit Function
'''                End If
'''            End If
'''NextInv:    N = PayN + 1
'''        Loop While N <= P.EOL
'''    End With
End Function
Function IgnoredFirm(ByVal Firm As String) As Boolean
'
' - IgnoredFirm(Firm)   - ���������� TRUE ��� ������������ ����
' 20.11.12

    IgnoredFirm = False
    If InStr(Firm, "������") <> 0 Then IgnoredFirm = True
End Function
Function SNhandl(Acc1C, PayN, StockSN, SNinSF, ContrADSK) As String
'
' SNhandl(Acc1C, PayN, StockSN) - ��������� SN � ��������� �����
'   7.5.12
'  13.5.12 - ��������� ���� SNinSF - SN ��� ���������� � SF
'  18.5.12 - ��������� SN ����������� � ���� SN+SN
'   7.6.12 - ������� �� SNhandl SN ��������� � SF

    If PayN = 0 Or StockSN = "" Then Exit Function

    Dim i As Integer
    Dim j As Integer
    Dim sN As String    ' SN � ������ ADSKfrSF
    Dim S As String
    Dim CtrADSK As String
    Dim AccStock As String
    
    SNhandl = "": SNinSF = "": ContrADSK = ""
    S = RemIgnoredSN(StockSN)
    If Len(S) < 12 Then Exit Function   ' ��� SN
    
    For i = 2 To EOL_ADSKfrSF
        sN = Sheets(ADSKfrSF).Cells(i, SFADSK_SN_COL)
        If sN <> "" Then
            If InStr(S, sN) <> 0 Then
                AccStock = Sheets(ADSKfrSF).Cells(i, SFADSK_ACC1C_COL)
                If AccStock <> Acc1C Then ContrADSK = "'" & AccStock & "':"
                CtrADSK = Sheets(ADSKfrSF).Cells(i, SFADSK_CONTRACT_COL)
                If InStr(S, CtrADSK) <> 0 Then S = Replace(S, CtrADSK, "")
                If SNinSF <> "" Then SNinSF = SNinSF & "+"
                SNinSF = SNinSF & sN
                S = Replace(S, sN, "")  ' �������� ��������� SN � Contract Autodesk
                If ContrADSK <> "" Then ContrADSK = ContrADSK & "+"
                If InStr(ContrADSK, CtrADSK) = 0 Then ContrADSK = ContrADSK & CtrADSK
                If IsNumeric(ContrADSK) Then ContrADSK = "'" & ContrADSK
                If S = "" Then Exit For
            End If
        End If
    Next i
    S = Compressor(S)
    If Len(S) >= 12 Then SNhandl = S
End Function
Sub testRemIgnoregSN()
    Dim t, Q, R
    t = RemIgnoredSN("456 765-67812345")
    Q = RemIgnoredSN("")
    R = RemIgnoredSN("456-5654323 ����� 456-���-567")
End Sub

Function RemIgnoredSN(Str) As String
'
' - RemIgnoredSN(S)   - �������� ������������ ���� � ������ � S
'   7.5.12
'   8.6.12 - ������� �������� (<12 ������) ������
'  10.6.12 - bug fix

    Dim Ch As String
    Dim S As String
    Dim i As Integer
    Dim W() As String
    
    S = Str
    For i = 1 To Len(S)
        Ch = Mid(S, i, 1)
        If (Ch > "9" Or Ch < "0") And Ch <> "-" Then Ch = " "
        Mid(S, i, 1) = Ch
    Next i
    W = Split(S, " ")
    S = ""
    For i = LBound(W) To UBound(W)
        If Len(W(i)) = 12 And Mid(W(i), 4, 1) = "-" Then
            If S <> "" Then S = S & "+"
            S = S & W(i)
        End If
    Next i
    RemIgnoredSN = S
End Function
Sub SNdedub()
'
' Sndedup() - ������������ SN ��������� �� ������ - ��������!
'   7.5.12

    Call SheetDedup("SN", 1)
End Sub
Function SN_ADSKbyStock(PayK, Acc, Dat, StockRec) As String
'
' - SN_ADSKbyStock(PayId, Acc, Dat, StockRec) - ���������� SN �������� ADSK
'         �� ��������� �����. � StockRec ������������ ���� "������" �� ������.
'         �������� �������� ������������ ���������� ������� PayId � ������ �� ������.
'         ���� �������� �� ������ - ���������� "".
'   24.5.12
'   18.6.12 - use TxDate

    Const MaxDeliveryDays = 70

    Dim StockDat As Date    ' = ���� �������� ������ �� ������
    Dim StockSN As String   ' = SN ������ Autodesk �� ������
    Dim StockSch As Integer ' = ����� ����� �� ������
    Dim i As Integer
    
    SN_ADSKbyStock = ""
    With Sheets(STOCK_SHEET)
        For i = 2 To EOL_Stock
            StockDat = TxDate(.Cells(i, STOCK_DATE_COL))
            If StockDat >= Dat And StockDat < Dat + MaxDeliveryDays Then
                If Acc = .Cells(i, STOCK_ACC_COL) Then
                    If .Cells(i, STOCK_GOOD_COL) = WE_GOODS_ADSK Then
                        StockSch = InvoiceN(.Cells(i, STOCK_INVOICE_COL))
                        If StockSch = PayInvByK(PayK) Then
                            StockSN = .Cells(i, STOCK_SF_SN_COL)
                            StockRec = .Cells(i, STOCK_CLIENT_COL)
                            If StockSN <> "" Then SN_ADSKbyStock = StockSN
                        End If
                    End If
                End If
            End If
        Next i
    End With
    
End Function
Function GetContrADSKfrStock(StockSN) As String
'
' - GetContrADSKfrStock (StockSN) - ���������� ��������� ADSK �� ��������� �����
'   18.5.12

    Dim i As Long
    Dim S As String
    
    GetContrADSKfrStock = ""
    
'!!!!!!!!!!! �������� !!!!!!!!!!!!!!!!!!!!!!!!
    If InStr(StockSN, "110000") = 0 Then Exit Function
    
    For i = 1 To Len(StockSN) - 12
        S = Mid(StockSN, i, 12)
        If IsNumeric(S) And InStr(S, "110000") = 1 Then
            GetContrADSKfrStock = S
            Exit Function
        End If
    Next i
End Function
