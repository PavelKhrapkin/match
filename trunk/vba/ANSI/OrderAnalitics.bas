Attribute VB_Name = "OrderAnalitics"
'---------------------------------------------------------------------------------
' OrderAnalitics  - ������ �� ����� ������� �.�������
'
' [*] OrderPass()     - ������ �� ����� ������� Orders
'
'   28.4.2012

Option Explicit

Sub OrderPass()
'
' [*] OrderPass()   - ������ �� ����� �������, ������������ ����� �������
'   28.4.12
'   16.6.12 - ���������, 5 ��������� �������

'------ INITIALIZATION AND LOCAL DECLARATION SECTION ---------------------
    Dim Inv1C As String     '= ����������� �� ����� ���� 1�
    Dim Dat As String       '���� ���� ����� CSD
    Dim PaidDat As String   '���� ���� ������ �������
    Dim Client As String    '���� ������ 1�
    Dim iOL As Integer      '= ����� ������ � �������
    Dim i1� As Integer      '= ����� ������ � ��������
    
    Dim i As Integer
    
    EOL_OrderList = ModStart(OrderList, "������ �� �������: ��������� � SF", True)
    
    CheckSheet OrderList, 1, OL_ORDERN_COL, OrderListStamp
    ClearSheet NewOrderList, Range("HDR_NewOrderList")
    
    EOL_SForders = EOL(SForders)
    
'---------------------- CODE SECTION -----------------------------------
    With Sheets(OrderList)
        '-- ��������� EOL_OrderList �������� �����
        EOL_OrderList = EOL_OrderList - OL_MIN_RESLINES
        Do While .Cells(EOL_OrderList, OL_ORDERN_COL) = ""
            EOL_OrderList = EOL_OrderList - 1
        Loop
         
        For i = 2 To EOL_OrderList
            Progress i / EOL_OrderList
            
If i >= 477 Then
i = i
End If
            Inv1C = .Cells(i, OL_INV_1C_COL)
            If Trim(Inv1C) = "" Then
                Inv1C = SeekInv(.Cells(i, OL_ORDERN_COL))
            End If
            Dat = .Cells(i, OL_CSDINVDAT_COL)
            If IsInv1C(Inv1C, Dat, i1�) Then
                With Sheets(PAY_SHEET)
                    PaidDat = .Cells(i1�, PAYDATE_COL)
                    Inv1C = .Cells(i1�, PAYINVOICE_COL)
                    Client = .Cells(i1�, PAYACC_COL)
                End With
            Else
                PaidDat = "": Inv1C = "": Client = ""
            End If
            .Cells(i, OL_PAIDDAT_COL) = PaidDat
            .Cells(i, OL_INV1C_COL) = Inv1C
            .Cells(i, OL_ACC1C_COL) = Client
            
'            OrderN = .Cells(i, OL_ORDERN_COL)
'            If Not IsOrderN(OrderN, iOL) Then
'                NewOrder (i)
'            End If
        Next i
    End With
'----------------------- SUMMARY SECTION -------------------------------
    ModEnd OrderList
End Sub
Function IsInv1C(Str, Dat, i1C) As Boolean
'
' - IsInv1C(Str, Dat, i1C)  - ���������� TRUE � ����� ������ � �������� 1�,
'                             ���� ���� �� Str ��������� � ������
'   16.6.12
    
    Const PO_DAYS = 50      ' ���������� ����� ���� �� ������� �� ������
    
    Dim Inv1C As String     '���� "����" ������� 1�
    Dim D As Date           '���� "���� ������� �����"
    Dim D_Min As Date
    Dim D_Max As Date
    
    IsInv1C = False
    If Not IsDate(Dat) Or Str = "" Then Exit Function
    
    D_Min = CDate(Dat) - PO_DAYS
    D_Max = CDate(Dat) + PO_DAYS
    
    For i1C = 2 To EOL_PaySheet
        Inv1C = Sheets(PAY_SHEET).Cells(i1C, PAYINVOICE_COL)
        If InStr(Inv1C, Str) <> 0 Then
            D = Sheets(PAY_SHEET).Cells(i1C, PAYDATE_COL)
            If D < D_Max And D > D_Min Then
                IsInv1C = True
                Exit Function
            End If
        End If
    Next i1C
End Function
Sub testInv1C()
    Dim i1C As Integer
    ModStart PAY_SHEET, "���� Inv1C"
    Call IsInv1C("��-278", "01.06.12", i1C)
End Sub

Function IsOrderN(OrderN, iOL) As Boolean
'
' ���� ����� OrderN ���� � SF, ���������� TRUE
' 28.4.12

    IsOrderN = False
    Dim i As Integer
    For i = 2 To EOL_SForders
        If OrderN = Sheets(SForders).Cells(i, SFORDERS_ORDERN_COL) Then
            IsOrderN = True
            Exit Function
        End If
    Next i
    Exit Function
End Function
Sub NewOrder(i)
'
' NewOrder - ������ ������ ������ � ���� NewOrderList ��� �������� � SF
'   28.4.12

    Dim j  ' ��������� ������ ����� P_Paid
    
    EOL_NewOrderList = EOL_NewOrderList + 1
    j = EOL_NewOrderList
    
    Dim invoice As Integer
    for i
    
    With Sheets(NewOrderList)
        .Cells(j, 1) = Sheets(OrderList).Cells(i, 6)            ' ����.���.
'        .Cells(j, 2) = DDMMYYYY(Sheets(1).Cells(i, 7))  ' ����
'        .Cells(j, 3) = Sheets(1).Cells(i, 8)            ' ����
'        .Cells(j, 4) = Dec(Sheets(1).Cells(i, 18))      ' ���� ���.
'        .Cells(j, 5) = Sheets(1).Cells(i, 19)           ' �����
'        .Cells(j, 6) = ContrId                          ' ContractId
'        .Cells(j, 7) = OppN                             ' OppN
    End With

End Sub
