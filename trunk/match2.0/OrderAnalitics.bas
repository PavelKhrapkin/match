Attribute VB_Name = "OrderAnalitics"
'---------------------------------------------------------------------------------
' OrderAnalitics  - проход по листу Заказов Д.Фролова
'
' S   GetInv1C(InvCol,PayN_Col, _       - находит номер строки в Платежах PayN со Счетом,
'         StrInvCol,DateCol,[Str2Inv])    найденным по строкам, содержащих Счет и по дате.
' [*] OrderPass()     - проход по листу Заказов Orders
'
'   4.9.2012

Option Explicit
Sub GetInv1C(InvCol As Integer, PayN_Col As Integer, _
    StrInvCol As Integer, DateCol As Integer, Optional Str2InvCol As Integer = 0)
'
'S GetInv1C(InvCol,PayN_Col, _       - находит номер строки в Платежах со Счетом 1С,
'      StrInvCol,DateCol,[Str2Inv])    найденным по строкам, содержащих Счет и по дате.
' ----- ПАРАМЕТРЫ, ЗАПИСЫВАЕМЫЕ В ТАБЛИЦУ ПРОЦЕССОВ -------
' 1.InvCol      - номер колонки в MyCol, куда вставляется найденная строка - Счет 1С
' 2.PayN_Col    - номер колонки в MyCol, куда всталяют найденный номер строки с Платежах 1С
' 3.StrInvCol   - номер колонки в Заказах - "Номер счета 1С"
' 4.DateCol     - номер колонки - привязка к Дате Счета 1С
' 5.[Str2InvCol]- альтернативная колонка с текстом Счета 1С
'-----------------------------------------------------------
'           * используется как Шаг для активных листов "Заказы" или "Склад"
'           * колонка Str2InvCol - возможная альтернативная строка со Счетом
'   26.8.12
' 4.9.12 - StepIn

    StepIn
    
    Dim DocTo As String ' имя входного Документа - отчета
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
            
            X = .Cells(i, DateCol): D = "0:0"
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

Sub OrderPass()
'
' [*] OrderPass()   - проход по листу Заказов, формирование Новых Заказов
'   28.4.12
'   18.6.12 - дописываю, 5 сиреневых колонок

'------ INITIALIZATION AND LOCAL DECLARATION SECTION ---------------------
    Dim Inv1C As String     '= извлекаемый из листа Счет 1С
    Dim Dat As String       'поле Дата Счета CSD
    Dim PaidDat As String   'поле Дата оплаты Платежа
    Dim Sale1C As String    'поле Продавец в 1С
    Dim Client As String    'поле Клиент 1С
    Dim iOL As Integer      '= номер строки в Заказах
    Dim i1С As Integer      '= номер строки в Платежах
    
    Dim i As Integer
    
    EOL_OrderList = ModStart(OrderList, "Проход по Заказам: Занесение в SF", True)
    
    CheckSheet OrderList, 1, OL_ORDERN_COL, OrderListStamp
    ClearSheet NewOrderList, Range("HDR_NewOrderList")
    
    EOL_SForders = EOL(SForders)
    
'---------------------- CODE SECTION -----------------------------------
    With Sheets(OrderList)
        '-- отодвинем EOL_OrderList выбросив пятку
        EOL_OrderList = EOL_OrderList - OL_MIN_RESLINES
        Do While .Cells(EOL_OrderList, OL_ORDERN_COL) = ""
            EOL_OrderList = EOL_OrderList - 1
        Loop
         
        For i = 2 To EOL_OrderList
            Progress i / EOL_OrderList
            
If i >= 113 Then
i = i
End If
            Inv1C = .Cells(i, OL_INV_1C_COL)
            If Trim(Inv1C) = "" Then
                Inv1C = SeekInv(.Cells(i, OL_ORDERN_COL))
            End If
            Dat = .Cells(i, OL_CSDINVDAT_COL)
            If IsInv1C(Inv1C, Dat, i1С) Then
                With Sheets(PAY_SHEET)
                    PaidDat = .Cells(i1С, PAYDATE_COL)
                    Inv1C = .Cells(i1С, PAYINVOICE_COL)
                    Sale1C = .Cells(i1С, PAYSALE_COL)
                    Client = .Cells(i1С, PAYACC_COL)
                    
                End With
            Else
                PaidDat = "": Inv1C = "": Client = "": Sale1C = ""
            End If
            .Cells(i, OL_PAIDDAT_COL) = PaidDat
            .Cells(i, OL_INV1C_COL) = Inv1C
            .Cells(i, OL_SALE1C_COL) = Sale1C
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
' - IsInv1C(Str, Dat, i1C)  - возвращает TRUE и номер строки в Платежах 1С,
'                             если Счет из Str распознан и найден
'   18.6.12
    
    Const PO_DAYS = 50      ' наибольшее число дней от Платежа до Заказа
    
    Dim Inv1C As String     'поле "Счет" Платежа 1С
    Dim D As Date           'поле "Дата прихода денег"
    Dim D_Min As Date
    Dim D_Max As Date
    
    IsInv1C = False
    If Not IsDate(Dat) Or Str = "" Then Exit Function
    
    D_Min = CDate(Dat) - PO_DAYS
    D_Max = CDate(Dat) + PO_DAYS
    
    For i1C = 2 To EOL_PaySheet
        Inv1C = DB_1C.Sheets(PAY_SHEET).Cells(i1C, PAYINVOICE_COL)
        If InStr(Inv1C, Str) <> 0 Then
            On Error Resume Next
            D = DB_1C.Sheets(PAY_SHEET).Cells(i1C, PAYDATE_COL)
            On Error GoTo 0
            If D < D_Max And D > D_Min Then
                IsInv1C = True
                Exit Function
            End If
        End If
    Next i1C
End Function
Sub testInv1C()
    Dim i1C As Integer
    ModStart PAY_SHEET, "Тест Inv1C"
    Call IsInv1C("Сч-278", "01.06.12", i1C)
End Sub

Function IsOrderN(OrderN, iOL) As Boolean
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
Sub WrOrderMyCol()
'
' S WrOrderMyCol()
' 4.9.12

    StepIn
    
    Dim R As TOCmatch   ' входной Документ "Заказы"
    Dim P As Range      ' форма с Адаптерами
    Dim i As Long, OrderLine As Long, Ind As Long
    Dim X As String, Y As String, IsErr As Boolean
    
    R = GetRep("Заказы")
    Set P = DB_STOCK.Sheets("Forms").Range("Orders_MyCol")
    
    With DB_STOCK.Sheets(R.SheetN)
        For OrderLine = 2 To R.EOL
            Progress OrderLine / R.EOL
            For i = 1 To R.MyCol
                Ind = P.Cells(4, i)
                If Ind > 0 Then
                    X = .Cells(OrderLine, Ind)
                    Y = Adapter(P.Cells(5, i), X, P.Cells(6, i), IsErr)
                    If IsErr Then
                        .Cells(OrderLine, i) = ""
                        .Cells(OrderLine, i).Interior.Color = rgbRed
                    Else
                        .Cells(OrderLine, i) = Y
                    End If
                End If
            Next i
        Next OrderLine
    End With
End Sub
Sub NewOrd()
'
' S NewOrder - запись Новых Заказов в лист NewOrderList для загрузки в SF
'   5.9.12

    StepIn

    Dim Ord As TOCmatch 'Заказы
    Dim i As Long
    
    Ord = GetRep(OrderList)
    
    With DB_MATCH.Sheets(NewOrderList)
        For i = 2 To Ord.EOL
            Progress i / Ord.EOL
            If .Cells(i, OL_IDSFORDER_COL) = "" And .Cells(i, OL_IDSF_COL) <> "" Then
                WrNewSheet NewOrderList, DB_STOCK.Sheets(OrderList), i
            End If
        Next i
    End With
End Sub
