Attribute VB_Name = "Checking"
'------------------------------------------------------
' Checking -- �������� ��������������� ���� Match
'
' (*) CheckPaySales()   - �������� �������� 1C-SF - ����� Ctlr/Sft/Q
'  - CheckSheet SheetN, R, C, Stamp - ��������, ��� SheetN ����������
' [*] CheckGoodType()   - ������ �� �������� � �������� ��� �� ������ ������������
' [*] ContractCheck()   - �������� ��������� ���������
'
'   19.2.2012
'   24.2.2012 - ���-�� ������ �� �������� � We
'   29.4.12 - �������� ��������� ���������

Option Explicit

Const WE_ERR_COL = 7    ' ������� � ������ �������������� �� ������� �������� � We
Const WE_GOOD_COL = 6   ' ������� ���������, ��������������� �������� � We

Sub CheckFofmOutput()
'
'   ���������� � ����� ����� CheckingForm
'   24.2.2012
    
    Dim S As Range
    Dim N, E
    
    N = 1
    CheckingForm.SalesList.RowSource = ""
    CheckingForm.SalesList.ColumnCount = 2
    For Each S In Sheets(We).Range("��������").Rows
        E = S.Cells(1, WE_ERR_COL)  ' ����� �������������� ��� ������� ��������
        If E > 0 Then
            CheckingForm.SalesList.AddItem S.Cells(1, 1).Value
            CheckingForm.SalesList.List(N - 1, 1) = E
            N = N + 1
        End If
    Next S
    CheckingForm.Show
End Sub
Sub CheckPaySales(SelectedSale)
'
' �������� �������� SelectedSale � �������� �� ��������.
' �������� � 1� - �������� ������� � SF?
' ���� SelectedSales = All - ��������� ��� ���� ���������
'   27.2.2012
    
    Dim i, j, OppN As Integer
    Dim AllSales As Boolean
    Dim K, SaleSF, Sale1C, ContrK, ErrMsg As String
    Dim S As Range

    Lines = ModStart(1, "�������� ������� ��������� 1� �� �������� SF." & _
        " �������� = " & SelectedSale, True) - 3
        
    AllSales = SelectedSale = "All"
'---- ���� ������ [��� ��������] - �������� �������� �������������� SF-1C
    If AllSales Then
        For Each S In Range("���_��������").Rows
            S.Cells(1, WE_ERR_COL) = 0
        Next S
    End If
' -------- ������ �� �������� -----------------
    For i = 2 To Lines
        Progress i / Lines
        If (AllSales Or Cells(i, 22) = SelectedSale) And _
            Cells(i, 1) = 1 And _
            Cells(i, 4) = 1 And _
            Trim(Cells(i, 6)) <> "" Then
                K = Cells(i, 5)             ' ��� �������
                Sale1C = Cells(i, 22)       ' �������� � 1�
                OppN = OppNbyPay(K)         ' ������ SF �� ���� �������
                SaleSF = OppOwner(K)
                ContrK = ContrCod(Cells(i, 25), Cells(i, 26))
                If Not ContrOK(OppN, ContrK) Then
                    ErrMsg = "�� ������� '" & Trim(Cells(i, 6)) & _
                        "' " & Cells(i, 9) & _
                        " ������� ������ � SF ������� '" & ContrK & "'"
                    Dim PayN, PayId, DogN
                    PayN = Cells(i, 2)
                    On Error Resume Next
                    PayId = WorksheetFunction.VLookup(PayN, _
                        Sheets(SF).Range("A:R"), 18, False)
                    DogN = WorksheetFunction.VLookup(PayN, _
                        Sheets(SF).Range("A:R"), 17, False)
                    On Error GoTo 0
'                    ErrMsg = ErrMsg & " � ������� <" & PayId & _
'                        "> �� ������� <" & DogN & ">"
                    If ContractId(ContrK) = "" Then
                        ErrMsg = ErrMsg & " -- ������ �������� ��� � SF!"
'                        ErrMsg = ErrMsg & " -- ��� ������ �� ������ " & OppN
'                        Dim DogOppN As String
'                        On Error Resume Next
'                        DogOppN = WorksheetFunction.VLookup(ContrK, _
'                            Sheets(SFD).Range("B:Q"), 16, False)
'                        On Error GoTo 0
'                        If DogOppN = "" Then
'                        End If
                    Else
                        ErrMsg = ErrMsg & " -- ����� ������� ������� � �������� " _
                            & OppN
                    End If
                    LogWr "     WARNING:  " & ErrMsg
                End If
                If Not IsSameTeam(Sale1C, SaleSF, OppN) Then
                    Call CheckSaleErr(Sale1C)
                    ErrMsg = "������ " & Trim(Cells(i, 6)) & vbTab & _
                            ", �������� 1�=" & Sale1C & vbTab & _
                            ", � � SF=" & SaleSF & ", ������ " & i & _
                            ", ������ SF=" & OppByPay(K)
                    LogWr ("�������������� SF � 1�! " & ErrMsg)
                End If
        End If
    Next i
'    ProgressForm.Hide
'    MsgBox ErrMsg, , "<CheckPaySales>" '    ErrMsg = "��������= " & SelectedSale & ". ����� ��������� �������� " & j
'    LogWr ErrMsg

    ModEnd 1
End Sub

Function CheckSaleErr(Sale1C) As Integer
'
' ������ � We.�������� �� ��������� ������� �� �������� Sale1C
' ���������� ���������� ������ �� �������� Sale1C, �������� ��� �� 1
' 23.2.2012

    Dim S As Range
    
    For Each S In Sheets(We).Range("���_��������").Rows
        If S.Cells(1, 1) = Sale1C Then
            CheckSaleErr = S.Cells(1, WE_ERR_COL) + 1
            S.Cells(1, WE_ERR_COL) = CheckSaleErr
            Exit For
        End If
    Next S
End Function
Function IsSameTeam(S1, S2, OppN) As Boolean
'
' ���������� True, ���� �������� S1 � S2 �������� ������ ��� ��� ���� � ��� ��.
' ������ SalesTeam � ������� �������� ����� [We], ��� � ���� SF ���������
'   17.2.2012
'   26.2.2012 - ���� ���� SF ��������� ������ ������� � We, ���������� � InStr
'   6.3.12 - ����������, ������ S1 � S2 ���������� � ������� ���������

    Dim S, T, Sales() As String  ' ������ ���������
    Dim X1, X2 As Range
    Dim i

    IsSameTeam = False
' �������� �� ������� ����������
    If InStr(S2, S1) <> 0 Then GoTo Found
' �������� �� We: ��� ��� �� ����� (���.3) ��� � ��� �� ������� (���.5)
    For Each X1 In Range("��������").Rows
        If InStr(S1, X1.Cells(1, 1)) <> 0 Then GoTo 1   ' ����� ������� �������� S1
    Next X1
    GoTo ErrorWe1
1:  For Each X2 In Range("��������").Rows
        If InStr(S2, X2.Cells(1, 1)) <> 0 Then GoTo 2   ' ����� ������� �������� S2
    Next X2
    GoTo ErrorWe2
2:  If X1.Cells(1, 1) = X2.Cells(1, 1) Then GoTo Found  ' ��� ���� � ��� �� ��������
    If InStr(X2.Cells(1, 5), X1.Cells(1, 1)) <> 0 Then GoTo Found ' ��� � ����� �������
    If InStr(X1.Cells(1, 5), X2.Cells(1, 1)) <> 0 Then GoTo Found ' ��� � ����� �������
' �������� �� ���� SF ���������
    On Error Resume Next        ' ������-�� (?) ���������� ������� � SFopp �����������
    S = WorksheetFunction.VLookup(Trim(OppN), _
        Sheets(SFopp).Range("B:K"), 10, False)
    On Error GoTo 0
    If S <> "" Then
        If InStr(S, X1.Cells(1, 1)) <> 0 Then GoTo Found
        If InStr(S, X2.Cells(1, 1)) <> 0 Then GoTo Found
    End If
    Exit Function
Found:
    IsSameTeam = True
    Exit Function
ErrorWe2:   S1 = S2
ErrorWe1:
    LogWr "ERROR: ������ � ��������� We: �� ������ �������� " & S1 & _
        " � OppN = '" & OppN & "'"
End Function
Function IsRightSale(Sale, GoodType) As Boolean
'
' ���������� True, ���� �������� Sale �������� � ������� ���� GoodType.
' ������� ������������ � ������� �������� ����� [We]
'   17.2.2012
'   18.3.12 - �� ������, ���� GoodType ="" - ����� � ���������

    Dim S, Goods() As String  ' ������ ��������� � �������
    Dim i
    Dim X As Range

    IsRightSale = False
    
    For Each X In Range("��������").Rows
        If InStr(Sale, X.Cells(1, 1)) <> 0 Then Exit For   ' ����� ������� �������� S1
    Next X

    On Error Resume Next
    S = WorksheetFunction.VLookup(X.Cells(1, 1), _
            Range("��������"), WE_GOOD_COL, False)
    On Error GoTo 0
    If S = "" Then
        If GoodType <> "" Then
            LogWr ("ERROR!  �������� ��������='" & Sale & "' � T����='" & GoodType) & "'"
        End If
        Exit Function
    End If
'==== ������ � Goods ������ ������� (Goods) �� ������� We, � �������� �������� ��������
    Goods = Split(S, ",")
    For i = 0 To UBound(Goods)
        If Trim(Goods(i)) = GoodType Then
            IsRightSale = True
            Exit Function
        End If
    Next i
End Function
Sub CheckSheet(SheetN, R, C, txt)
'
' ��������, ��� � ������ ����� SheetN ������������� ����� ����� Txt
'   26/1/2012

    Sheets(SheetN).Select
    If Cells(R, C) <> txt Then
        Cells(R, C).Activate
        MsgBox "������������ ���� " & SheetN, vbCritical, "ERROR!"
        Stop
    End If
End Sub

Sub CheckGoodType()
'
' ������ �� �������� � ��������, ��� �� ���� ������� ������������
'   12.3.12

    Dim i, X As Integer
    Dim T As String
    
'    profileGlobal = 0
        
    Lines = ModStart(1, "CheckGoodType", True)
    
    X = 0
    For i = 2 To Lines - 3
       Call Progress(i / Lines)
'       profileGlobal = getPoint()
'       totalTime = getPoint() - profileGlobal      ' end (may be invoked repeatedly)
       If Sheets(1).Cells(i, 1) = 1 And _
          Trim(Sheets(1).Cells(i, 6)) <> "" Then
            T = GoodType(Sheets(1).Cells(i, 19))
            If T = "" Then X = X + 1
        End If
    Next i
    MsgBox "�� ���������� ������ � " & X & " ��������"
    ModEnd 1
End Sub
Sub ContractCheck()
'
' [*] ContractCheck()   - �������� ��������� ���������
'   29.4.12

    Dim MSG, DogSFstat, ContrK, ContrId As String
    Dim i, DogPaid

    Lines = ModStart(DOG_SHEET, "ContractCheck: ������ �� ���������", True) - DOGRES
    
    CheckSheet DOG_SHEET, 1, 10, Stamp1Cdog1
    CheckSheet SFD, EOL_SFD + 2, 3, SFcontrRepName
    CheckSheet SFopp, EOL_SFopp + 2, 1, SFoppRepName
'---------- ������ �� ��������� ------------------------
    
    For i = 2 To Lines
        Progress i / Lines
        If ExRespond = False Then Exit For
        With Sheets(DOG_SHEET)
            DogSFstat = .Cells(i, DOGSFSTAT_COL)    ' ������ �������� � SF
            DogPaid = .Cells(i, DOGPIAID1C_COL)     ' ���� 1 - ������� ������� � 1�
            ContrK = .Cells(i, DOGCOD_COL)
            ContrId = ContractId(ContrK)

            MSG = ""
            If DogPaid = "1" And DogSFstat <> "������" Then
                MSG = "�������! ������� " & ContrK & " (" & ContrId & ") �������, " _
                    & " � � SF ��� ������ '" & DogSFstat & "'"
                LogWr MSG
            End If
        End With
    Next i
    
    ModEnd DOG_SHEET
End Sub
