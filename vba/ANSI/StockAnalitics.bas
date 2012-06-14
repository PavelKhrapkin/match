Attribute VB_Name = "StockAnalitics"
'---------------------------------------------------------------------------------
' StockAnalitics  - ������ ���� �� ������
'
' [*] StockHandling()   - ������ �� ����� "�����", ����� ��������� Autodesk
'  -  FindAcc1C(Client, Acc1C) - ����� ����� 1� ���������� � Client ��� Acc1C
'  -  SNhandl(Acc1C, PayN, StockSN) - ��������� SN � ��������� �����
'  ?  RemIgnoredSN(S)   - �������� ������������ ���� � ������ � S �� SN
' (*) Sndedup() - ������������ SN ��������� �� ������ - ��������!
'  -  SN_ADSKbyStock(PayId, Acc, Dat) - ���������� SN �������� ADSK �� ������

'   10.6.2012

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
    Dim Good, T As String   ' ����� (������������) � ��� ������
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
If i >= 866 Then
i = i
End If
            Client = .Cells(i, STOCK_CLIENT_COL)
            If Client = .Cells(i - 1, STOCK_CLIENT_COL) Then
                SameClient = True
            Else
                SameClient = False
                Acc1C = FindAcc(Client)                 ' ����� �����������
                PayN = FindPayN(Client, Acc1C)          ' ����� �������
            End If
            .Cells(i, STOCK_ACC_COL) = Acc1C
            If PayN > 0 And PayN <= EOL_PaySheet Then
                .Cells(i, STOCK_INVOICE_COL) = _
                        Sheets(PAY_SHEET).Cells(PayN, PAYINVOICE_COL)
                Good = Sheets(PAY_SHEET).Cells(PayN, PAYGOOD_COL)
                T = GoodType(Good)              ' ��� ������ �� �����
                .Cells(i, STOCK_GOOD_COL) = T
                If T = WE_GOODS_ADSK Then
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
Function FindPayN(Client, Acc1C) As Integer
'
' FindAcc1C(Client) - ����� ������� �� �����1� ����������� � Client ��� Acc1C
'   7.5.12
'  23.5.12 -�� ����� � Acc1C ����� ���� ��������� ����������� ����������� "+"
'           FindPayN ���� ���� �� ���� � ��������� � Acc1C ������ ��������

    FindPayN = 0
    If Acc1C = "" Or Acc1C = "*" Then Exit Function
    
'----- ��������� ����� �� ������ � ��������� ����� Client ----
    Dim Invoice As String
    Dim SeekInv As String
    Dim Sch As String
    Dim i As Integer

    Dim StWord() As String
    StWord = split(LCase(Client), " ")
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
    SeekInv = Sch
    
'----- ����� ����� � �������� -------
    Dim Inv As String, Acc As String
    Dim j As Integer
    
    For i = 2 To EOL_PaySheet
        Acc = Sheets(PAY_SHEET).Cells(i, PAYACC_COL)
        If InStr(Acc1C, Acc) <> 0 Then
            Invoice = Sheets(PAY_SHEET).Cells(i, PAYINVOICE_COL)
            If InStr(Invoice, SeekInv) <> 0 Then
                For j = 4 To 7 ' ��������� ������ ����� ���� "��-123 ..."
                    If Mid(Invoice, j, 1) = " " Then
                        Inv = Mid(Invoice, 4, j - 4)
                        Exit For
                    End If
                Next j
                If Inv = SeekInv Then
                    FindPayN = i
                    Acc1C = Acc
                    Exit Function
                End If
            End If
        End If
    Next i
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
    Dim T, Q, R
    T = RemIgnoredSN("456 765-67812345")
    Q = RemIgnoredSN("")
    R = RemIgnoredSN("456-5654323 ����� 456-���-567")
End Sub

Function RemIgnoredSN(Str) As String
'
' - RemIgnoredSN(S)   - �������� ������������ ���� � ������ � S
'   7.5.12
'   8.6.12 - ������� �������� (<12 ������) ������
'  10.6.12 - bug fix -- Replace 1 ���!

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
    For i = LBound(W) To UBound(W)
        If Len(W(i)) < 12 Then S = Replace(S, W(i), " ", , 1)
    Next i
    RemIgnoredSN = Compressor(S)
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

    Const MaxDeliveryDays = 70

    Dim DatStr As String    '���� - ���� �������� ������ �� ������
    Dim StockDat As Date    ' = ���� �������� ������ �� ������
    Dim StockSN As String   ' = SN ������ Autodesk �� ������
    Dim StockSch As Integer ' = ����� ����� �� ������
    Dim i As Integer
    
    SN_ADSKbyStock = ""
    With Sheets(STOCK_SHEET)
        For i = 2 To EOL_Stock
            DatStr = .Cells(i, STOCK_DATE_COL)
            StockDat = "1.1.2000"
            If IsDate(DatStr) Then StockDat = DatStr
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
