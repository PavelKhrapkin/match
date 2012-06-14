Attribute VB_Name = "OrderAnalitics"
'---------------------------------------------------------------------------------
' OrderAnalitics  - ������ �� ����� ������� �.�������
'
' [*] OrderPass()     - ������ �� ����� ������� Orders
'
'   28.4.2012

Option Explicit

Const EOLSTART = 100        ' �������� ���������� ������ �����

Sub OrderPass()
'
' [*] OrderPass()   - ������ �� ����� �������, ������������ ����� �������
'   28.4.12

    Dim i As Integer
    
    ModStart OrderList, "������ �� �������: ��������� � SF", True
    
    CheckSheet OrderList, 2, 3, OrderListStamp
    ClearSheet NewOrderList, Range("HDR_NewOrderList")
    EOL_OrderList = EOL_Order(Lines)
    EOL_SForders = EOL(SForders)
    
    With Sheets(OrderList)
        For i = 4 To EOL_OrderList
            Progress i / EOL_OrderList
            OrderN = .Cells(i, OL_ORDERN_COL)
            If Not IsOrderN(OrderN) Then
                NewOrder (i)
            End If
        Next i
    End With

    ModEnd OrderList
End Sub
Function IsOrderN(OrderN) As Boolean
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
Function EOL_Order(Lines)
'
' ����� ��������� �������� ������ � ����� �������
'   28.4.12

    Dim i As Integer
    
    i = Lines
    Do While Sheets(OrderList).Cells(i, OL_ORDER_COL) <> ""
        i = i + 1
    Loop
        
    EOL_Order = i - 1
End Function
