Attribute VB_Name = "OrderAnalitics"
'---------------------------------------------------------------------------------
' OrderAnalitics  - проход по листу Заказов Д.Фролова
'
' [*] OrderPass()     - проход по листу Заказов Orders
'
'   28.4.2012

Option Explicit

Const EOLSTART = 100        ' заведомо избуточный размер пятки

Sub OrderPass()
'
' [*] OrderPass()   - проход по листу Заказов, формирование Новых Заказов
'   28.4.12

    Dim i As Integer
    
    ModStart OrderList, "Проход по Заказам: Занесение в SF", True
    
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
' если Заказ OrderN есть в SF, возвращает TRUE
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
' NewOrder - запись Нового Заказа в лист NewOrderList для загрузки в SF
'   28.4.12

    Dim j  ' последняя строка листа P_Paid
    
    EOL_NewOrderList = EOL_NewOrderList + 1
    j = EOL_NewOrderList
    
    Dim invoice As Integer
    for i
    
    With Sheets(NewOrderList)
        .Cells(j, 1) = Sheets(OrderList).Cells(i, 6)            ' Плат.док.
'        .Cells(j, 2) = DDMMYYYY(Sheets(1).Cells(i, 7))  ' Дата
'        .Cells(j, 3) = Sheets(1).Cells(i, 8)            ' Счет
'        .Cells(j, 4) = Dec(Sheets(1).Cells(i, 18))      ' Итог руб.
'        .Cells(j, 5) = Sheets(1).Cells(i, 19)           ' Товар
'        .Cells(j, 6) = ContrId                          ' ContractId
'        .Cells(j, 7) = OppN                             ' OppN
    End With

End Sub
Function EOL_Order(Lines)
'
' поиск последней значащей строки в файле Заказов
'   28.4.12

    Dim i As Integer
    
    i = Lines
    Do While Sheets(OrderList).Cells(i, OL_ORDER_COL) <> ""
        i = i + 1
    Loop
        
    EOL_Order = i - 1
End Function
