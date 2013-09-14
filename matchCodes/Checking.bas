Attribute VB_Name = "Checking"
'------------------------------------------------------
' Checking -- проверка консистентности базы Match
'
'[*] CheckFofmOutput()  - Подготовка и вывод формы CheckingForm
'[*] CheckPaySales()   - проверка Продавца 1C-SF - вызов Ctlr/Sft/Q
' -  IsSameTeam(S1, S2, Optional OppN = "")  - возвращает TRUE, если Продавцы S1 и S2
'                                              работали вместе или это один и тот же.
' -  IsRightSale(Sale, GoodType) - Возвращает True, если Продавец Sale
'                                  работает с Товаром типа GoodType.
'[*] CheckGoodType()   - проход по Платежам и проверка все ли товары распознаются
'[*] ContractCheck()   - Проверка состояния Договоров
' S  CheckRepDate(Rep1,[Rep2],[Rep3],[Rep4],[Rep5])  - проверка дат Документов
' -  RepDateSub(Rep, Name, Dat)  - возвращает TRUE, если Документ Rep актуален
'                         иначе выводит сообщение о необходимости перезагрузки
'
' 14.9.13
' --- история модуля ----
'   19.2.2012
'   24.2.2012 - кол-во ошибок по Продавцу в We
'   29.4.12 - проверка состояния Договоров
'   18.6.12 - Optional OppN в InSameTeam для использования в Заказах
'   14.9.13 - CheckRepDate - проверка дат Документов

Option Explicit

Const WE_ERR_COL = 7    ' колонка с числом Несоответствий по данному Продавцу в We
Const WE_GOOD_COL = 6   ' колонка Продукция, соответствующая Продавцу в We

Sub CheckFofmOutput()
'
'[*] CheckFofmOutput()  - Подготовка и вывод формы CheckingForm
'   24.2.2012
    
    Dim S As Range
    Dim N, E
    
    N = 1
    CheckingForm.SalesList.RowSource = ""
    CheckingForm.SalesList.ColumnCount = 2
    For Each S In Sheets(We).Range("Продавцы").Rows
        E = S.Cells(1, WE_ERR_COL)  ' число несоответствий для данного Продавца
        If E > 0 Then
            CheckingForm.SalesList.AddItem S.Cells(1, 1).Value
            CheckingForm.SalesList.list(N - 1, 1) = E
            N = N + 1
        End If
    Next S
    CheckingForm.Show
End Sub
Sub CheckPaySales(SelectedSale)
'
' Проверка Продавца SelectedSale с проходом по Платежам.
' Продавец в 1С - владелец Проекта в SF?
' Если SelectedSales = All - проверяем для всех Продавцов
'   27.2.2012
'   05.10.12 - адресация DB_MATCH.Sheets(We).
    
    Dim i, j, OppN As Integer
    Dim AllSales As Boolean
    Dim k, SaleSF, Sale1C, ContrK, ErrMsg As String
    Dim S As Range

    Lines = ModStart(1, "Проверка фамилий Продавцов 1С по Проектам SF." & _
        " Продавец = " & SelectedSale, True) - 3
        
    AllSales = SelectedSale = "All"
'---- если кнопка [Все Продавцы] - обнуляем счетчики Несоответствий SF-1C
    If AllSales Then
        For Each S In DB_MATCH.Sheets(We).Range("ВСЕ_ПРОДАВЦЫ").Rows
            S.Cells(1, WE_ERR_COL) = 0
        Next S
    End If
' -------- проход по Платежам -----------------
    For i = 2 To Lines
        Progress i / Lines
        If (AllSales Or Cells(i, 22) = SelectedSale) And _
            Cells(i, 1) = 1 And _
            Cells(i, 4) = 1 And _
            Trim(Cells(i, 6)) <> "" Then
                k = Cells(i, 5)             ' код Платежа
                Sale1C = Cells(i, 22)       ' Продавец в 1С
                OppN = OppNbyPay(k)         ' Проект SF по Коду Платежа
                SaleSF = OppOwner(k)
                ContrK = ContrCod(Cells(i, 25), Cells(i, 26))
                If Not ContrOK(OppN, ContrK) Then
                    ErrMsg = "По Платежу '" & Trim(Cells(i, 6)) & _
                        "' " & Cells(i, 9) & _
                        " неверно связан в SF Договор '" & ContrK & "'"
                    Dim PayN, PayId, DogN
                    PayN = Cells(i, 2)
                    On Error Resume Next
                    PayId = WorksheetFunction.VLookup(PayN, _
                        DB_SFDC.Sheets(SF).Range("A:R"), 18, False)
                    DogN = WorksheetFunction.VLookup(PayN, _
                        DB_SFDC.Sheets(SF).Range("A:R"), 17, False)
                    On Error GoTo 0
'                    ErrMsg = ErrMsg & " и Платежа <" & PayId & _
'                        "> на Договор <" & DogN & ">"
                    If ContractId(ContrK) = "" Then
                        ErrMsg = ErrMsg & " -- такого Договора нет в SF!"
'                        ErrMsg = ErrMsg & " -- нет ссылки на Проект " & OppN
'                        Dim DogOppN As String
'                        On Error Resume Next
'                        DogOppN = WorksheetFunction.VLookup(ContrK, _
'                            Sheets(SFD).Range("B:Q"), 16, False)
'                        On Error GoTo 0
'                        If DogOppN = "" Then
'                        End If
                    Else
                        ErrMsg = ErrMsg & " -- можно связать Договор с Проектом " _
                            & OppN
                    End If
                    LogWr "     WARNING:  " & ErrMsg
                End If
                If Not IsSameTeam(Sale1C, SaleSF, OppN) Then
                    Call CheckSaleErr(Sale1C)
                    ErrMsg = "Платеж " & Trim(Cells(i, 6)) & vbTab & _
                            ", Продавец 1С=" & Sale1C & vbTab & _
                            ", а в SF=" & SaleSF & ", строка " & i & _
                            ", Проект SF=" & OppByPay(k)
                    LogWr ("Несоответствие SF и 1С! " & ErrMsg)
                End If
        End If
    Next i
'    ProgressForm.Hide
'    MsgBox ErrMsg, , "<CheckPaySales>" '    ErrMsg = "Продавец= " & SelectedSale & ". Всего ошибочных Платежей " & j
'    LogWr ErrMsg

    ModEnd 1
End Sub

Function CheckSaleErr(Sale1C) As Integer
'
' запись в We.Продавцы об ошибочном Платеже по Продавцу Sale1C
' возвращает количество ошибок по Продавцу Sale1C, увеличив его на 1
' 23.2.2012

    Dim S As Range
    
    For Each S In DB_MATCH.Sheets(We).Range("ВСЕ_ПРОДАВЦЫ").Rows
        If S.Cells(1, 1) = Sale1C Then
            CheckSaleErr = S.Cells(1, WE_ERR_COL) + 1
            S.Cells(1, WE_ERR_COL) = CheckSaleErr
            Exit For
        End If
    Next S
End Function
Function IsSameTeam(S1, S2, Optional OppN = "") As Boolean
'
' - IsSameTeam(S1, S2, Optional OppN = "")  - возвращает TRUE, если Продавцы S1 и S2
'                                             работали вместе или это один и тот же.
' Состав SalesTeam в массиве Продавцы листа [We], или в поле SF Компаньон
'   17.2.2012
'   26.2.2012 - учет поля SF Компаньон помимо профиля в We, переписано с InStr
'   6.3.12 - переписано, теперь S1 и S2 приводятся к фамилии продавцов
'   18.6.12 - Optional OppN для использования в Заказах
'   5.10.12 - адресация DB_MATCH.Sheets(We)

    Dim S, t, Sales() As String  ' массив Продавцов
    Dim X1, X2 As Range
    Dim i

    IsSameTeam = False
' проверка по входным параметрам
    If InStr(S2, S1) <> 0 Then GoTo Found
' проверка по We: это тот же самый (кол.3) или в той же команде (кол.5)
    For Each X1 In DB_MATCH.Sheets(We).Range("Продавцы").Rows
        If InStr(S1, X1.Cells(1, 1)) <> 0 Then GoTo 1   ' поиск фамилии Продавца S1
    Next X1
    GoTo ErrorWe1
1:  For Each X2 In DB_MATCH.Sheets(We).Range("Продавцы").Rows
        If InStr(S2, X2.Cells(1, 1)) <> 0 Then GoTo 2   ' поиск фамилии Продавца S2
    Next X2
    GoTo ErrorWe2
2:  If X1.Cells(1, 1) = X2.Cells(1, 1) Then GoTo Found  ' это один и тот же Продавец
    If InStr(X2.Cells(1, 5), X1.Cells(1, 1)) <> 0 Then GoTo Found ' они в одной команде
    If InStr(X1.Cells(1, 5), X2.Cells(1, 1)) <> 0 Then GoTo Found ' они в одной команде
' проверка по полю SF Компаньон
    If OppN = "" Then Exit Function
    On Error Resume Next        ' почему-то (?) некоторыые Проекты в SFopp отсутствуют
    S = WorksheetFunction.VLookup(Trim(OppN), _
        DB_SFDC.Sheets(SFopp).Range("B:K"), 10, False)
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
    LogWr "ERROR: Ошибка в структуре We: не найден Продавец " & S1 & _
        " в OppN = '" & OppN & "'"
End Function
Function IsRightSale(Sale, GoodType) As Boolean
'
' - IsRightSale(Sale, GoodType) - Возвращает True, если Продавец Sale
'                                 работает с Товаром типа GoodType.
' Таблица соответствия в массиве Продавцы листа [We]
'   17.2.2012
'   18.3.12 - не ошибка, если GoodType ="" - часто в Договорах
'   5.10.12 - адресация We в DB_MATCH

    Dim S, Goods() As String  ' массив Продавцов и товаров
    Dim i
    Dim X As Range

    IsRightSale = False
    
    For Each X In DB_MATCH.Sheets(We).Range("Продавцы").Rows
        If InStr(Sale, X.Cells(1, 1)) <> 0 Then Exit For   ' поиск фамилии Продавца S1
    Next X

    On Error Resume Next
    S = WorksheetFunction.VLookup(X.Cells(1, 1), _
            DB_MATCH.Sheets(We).Range("Продавцы"), WE_GOOD_COL, False)
    On Error GoTo 0
    If S = "" Then
        If GoodType <> "" Then
            LogWr ("ERROR!  Странный Продавец='" & Sale & "' и Tовар='" & GoodType) & "'"
        End If
        Exit Function
    End If
'==== разбор в Goods список товаров (Goods) из таблицы We, с которыми работает Продавец
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
' проверка, что в ячейке листа SheetN действительно лежит штамп Txt
'   26/1/2012

    Sheets(SheetN).Select
    If Cells(R, C) <> txt Then
        Cells(R, C).Activate
        MsgBox "Неправильный лист " & SheetN, vbCritical, "ERROR!"
        Stop
    End If
End Sub
Sub CheckGoodType()
'
' Проход по Платежам и проверка, все ли типы товаров распознаются
'   12.3.12

    Dim i, X As Integer
    Dim t As String
    
'    profileGlobal = 0
        
    Lines = ModStart(1, "CheckGoodType", True)
    
    X = 0
    For i = 2 To Lines - 3
       Call Progress(i / Lines)
'       profileGlobal = getPoint()
'       totalTime = getPoint() - profileGlobal      ' end (may be invoked repeatedly)
       If Sheets(1).Cells(i, 1) = 1 And _
          Trim(Sheets(1).Cells(i, 6)) <> "" Then
            t = GoodType(Sheets(1).Cells(i, 19))
            If t = "" Then X = X + 1
        End If
    Next i
    MsgBox "Не распознаны Товары в " & X & " Платежах"
    ModEnd 1
End Sub
Sub ContractCheck()
'
' [*] ContractCheck()   - Проверка состояния Договоров
'   29.4.12

    Dim msg, DogSFstat, ContrK, ContrId As String
    Dim i, DogPaid

    Lines = ModStart(DOG_SHEET, "ContractCheck: Проход по Договорам", True) - DOGRES
    
    CheckSheet DOG_SHEET, 1, 10, Stamp1Cdog1
    CheckSheet SFD, EOL_SFD + 2, 3, SFcontrRepName
    CheckSheet SFopp, EOL_SFopp + 2, 1, SFoppRepName
'---------- проход по Договорам ------------------------
    
    For i = 2 To Lines
        Progress i / Lines
        If ExRespond = False Then Exit For
        With Sheets(DOG_SHEET)
            DogSFstat = .Cells(i, DOGSFSTAT_COL)    ' статус Договора в SF
            DogPaid = .Cells(i, DOGPIAID1C_COL)     ' если 1 - Договор оплачен в 1С
            ContrK = .Cells(i, DOGCOD_COL)
            ContrId = ContractId(ContrK)

            If DogPaid = "1" And DogSFstat <> DOG_STAT_CLOSED Then
                msg = "ЗАКРЫТЬ! Договор " & ContrK & " (" & ContrId & ") оплачен, "
                Select Case DogSFstat
                Case DOG_STAT_OPEN:
                    msg = msg & " в SF Открыт, "
                    LoWr msg
                Case Else
                    MsgBox "Странный статус Договора " & ContrK & " '" & DogStat & "'"
                    Stop
                    ExRespond = False
                End Select
'                    & " а в SF его статус '" & DogSFstat & "'"
'                LogWr MSG
            End If
        End With
    Next i
    
    ModEnd DOG_SHEET
End Sub

Sub CheckRepDate(ByVal Rep1 As String, _
    Optional Rep2 As String = "", Optional Rep3 As String = "", _
    Optional Rep4 As String = "", Optional Rep5 As String = "")
'
' S CheckRepDate(Rep1,[Rep2],[Rep3],[Rep4],[Rep5])  - проверка дат Документов
'           взаимодействующих с загружаемым как самый левый в списке Шага
'
' 30.8.13
' 14.9.13 - изменен интерфейс RepDateSub, перенос в модуль Checking

    Dim ActTOC As TOCmatch
    
    StepIn
    
    ActTOC = GetRep(ActiveSheet.Name)   'на случай, если AcualSheet изменится
                                        '.. после StepIn
    With ActTOC
        RepDateSub Rep1, .Name
        RepDateSub Rep2, .Name
        RepDateSub Rep3, .Name
        RepDateSub Rep4, .Name
        RepDateSub Rep5, .Name
    End With
End Sub
Sub RepDateSub(Rep As String, Name As String)
'
' - RepDateSub(Rep, Name, Dat)  - возвращает TRUE, если Документ Rep актуален,
'                       то есть создан Документ Rep создан ПОЗЖЕ него.
'                       иначе выводит сообщение о необходимости перезагрузки
' 30.8.13
' 14.9.13 - изменен интерфейс, перенос в модуль Checking

    Dim LocalTOC As TOCmatch, ActTOC As TOCmatch
    
    If Rep = "" Then GoTo Ex
    
    LocalTOC = GetRep(Rep)
    ActTOC = GetRep(Name)
    
    If LocalTOC.CreateDat < ActTOC.CreateDat Then GoTo Er
      
Ex: Exit Sub
Er: ErrMsg FATAL_ERR, "Необходимо загрузить заново '" & Rep & "' за " _
        & LocalTOC.CreateDat & vbCrLf & "Он устарел относительно '" _
        & Name & "' за " & ActTOC.CreateDat
End Sub

