Attribute VB_Name = "StockAnalitics"
'---------------------------------------------------------------------------------
' StockAnalitics  - ������ ���� �� ������
'
' S   GetInv1C(InvCol,PayN_Col, _       - ������� ����� ������ � �������� PayN �� ������,
'         StrInvCol,DateCol,[Str2Inv])    ��������� �� �������, ���������� ���� � �� ����.
' [*] StockHandling()   - ������ �� ����� "�����", ����� ��������� Autodesk
'  -  FindAcc1C(Client, Acc1C) - ����� ����� 1� ���������� � Client ��� Acc1C
'  -  SeekInv(Str) - ��������� ����� � ��������� ������ Str
'  -  SNhandl(Acc1C, PayN, StockSN) - ��������� SN � ��������� �����
'  ?  RemIgnoredSN(S)   - �������� ������������ ���� � ������ � S �� SN
' (*) Sndedup() - ������������ SN ��������� �� ������ - ��������!
'  -  SN_ADSKbyStock(PayId, Acc, Dat) - ���������� SN �������� ADSK �� ������

'   26.8.2012

Option Explicit
Sub GetInv1C(InvCol As Integer, PayN_Col As Integer, _
    StrInvCol As Integer, DateCol As Integer, Optional Str2InvCol As Integer = 0)
'
'S GetInv1C(InvCol,PayN_Col, _       - ������� ����� ������ � �������� �� ������ 1�,
'      StrInvCol,DateCol,[Str2Inv])    ��������� �� �������, ���������� ���� � �� ����.
' ----- ���������, ������������ � ������� ��������� -------
' 1.InvCol      - ����� ������� � MyCol, ���� ����������� ��������� ������ - ���� 1�
' 2.PayN_Col    - ����� ������� � MyCol, ���� �������� ��������� ����� ������ � �������� 1�
' 3.StrInvCol   - ����� ������� � ������� - "����� ����� 1�"
' 4.DateCol     - ����� ������� - �������� � ���� ����� 1�
' 5.[Str2InvCol]- �������������� ������� � ������� ����� 1�
'-----------------------------------------------------------
'           * ������������ ��� ��� ��� �������� ������ "������" ��� "�����"
'           * ������� Str2InvCol - ��������� �������������� ������ �� ������
'   26.8.12

    Dim DocTo As String ' ��� �������� ��������� - ������
    Dim RepSF As TOCmatch, RepSForder As TOCmatch, RepP As TOCmatch, RepTo As TOCmatch
    Dim Str As String, Inv As String, Str2 As String, Inv2 As String, D As Date
    Dim i As Integer, i1C As Integer
    Dim X
    
    DocTo = ActiveSheet.Name
    RepTo = GetRep(DocTo)
    RepP = GetRep(PAY_SHEET)
    EOL_PaySheet = RepP.EOL
    

    With Workbooks(RepTo.RepFile).Sheets(RepTo.SheetN)
        .Activate
        For i = 2 To RepTo.EOL
            Progress i / RepTo.EOL
            
            Str = SeekInv(.Cells(i, StrInvCol))
            
            X = .Cells(i, DateCol): D = Now
            If IsDate(X) Then D = X
            
            Str2 = ""
            If Str2InvCol > 0 Then Str2 = SeekInv(.Cells(i, Str2InvCol))
            
            If Str = Str2 Then
            ElseIf Str = "" Or Str2 = "" Then
                Str = Str & Str2
            Else
                Str = Str & "+" & Str2
            End If
            .Cells(i, InvCol) = Str
            
            If IsInv1C(Str, D, i1C) Then .Cells(i, PayN_Col) = i1C
        Next i
    End With

End Sub

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
    Dim good, t As String   ' ����� (������������) � ��� ������
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
                good = Sheets(PAY_SHEET).Cells(PayN, PAYGOOD_COL)
                t = GoodType(good)              ' ��� ������ �� �����
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
        AccInv = Compressor(Sheets(PAY_SHEET).Cells(i1C, PAYACC_COL))
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
'        Acc = Sheets(PAY_SHEET).Cells(i, PAYACC_COL)
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
    StWord = split(S, " ")
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
    Dim A(1 To 6) As String
    A(1) = SeekInv("����� ��� ""���-94"" ��-267 �� 07.10.11 ��������� ")
    A(2) = SeekInv("����� ��� ""���-94"" �-267 �� 07.10.11 ��������� ")
    A(3) = SeekInv("����� ��� ""���-94"" � -267 �� 07.10.11 ��������� ")
    A(4) = SeekInv("����� ��� ""���-94"" ��- 267 �� 07.10.11 ��������� ")
    A(5) = SeekInv("����� ��� ""���-94"" �� - 267 �� 07.10.11 ��������� ")  '!!! �� �����������!!!
End Sub
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
    Dim SN As String    ' SN � ������ ADSKfrSF
    Dim S As String
    Dim CtrADSK As String
    Dim AccStock As String
    
    SNhandl = "": SNinSF = "": ContrADSK = ""
    S = RemIgnoredSN(StockSN)
    If Len(S) < 12 Then Exit Function   ' ��� SN
    
    For i = 2 To EOL_ADSKfrSF
        SN = Sheets(ADSKfrSF).Cells(i, SFADSK_SN_COL)
        If SN <> "" Then
            If InStr(S, SN) <> 0 Then
                AccStock = Sheets(ADSKfrSF).Cells(i, SFADSK_ACC1C_COL)
                If AccStock <> Acc1C Then ContrADSK = "'" & AccStock & "':"
                CtrADSK = Sheets(ADSKfrSF).Cells(i, SFADSK_CONTRACT_COL)
                If InStr(S, CtrADSK) <> 0 Then S = Replace(S, CtrADSK, "")
                If SNinSF <> "" Then SNinSF = SNinSF & "+"
                SNinSF = SNinSF & SN
                S = Replace(S, SN, "")  ' �������� ��������� SN � Contract Autodesk
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
    W = split(S, " ")
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
