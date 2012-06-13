Attribute VB_Name = "PaidAnalitics"
'---------------------------------------------------------------------------
' PaidAnalitics -- Макросы для анализа оплаченных Проектов
'
' * PaidHandling()  - проход по листу Платежей 1С, занесение в SF
' - GoodType(Good)              - возвращает строку - тип товара Good
' - IsSubscription(Good, GT)    - возвращает True, если товар - подписка
'
'   20.5.2012

Option Explicit

Const MinNewOpp = 120000
Const MinNewOppDialog = 200000  ' надо эти параметры запрашивать в форме

Sub PaidHandling()
'
' Проход по Платежам и поиск новых Платежей по незакрытым Проектам
'   7.3.2012
'  11.3.12 - просмотр существующих Платежей для связи с Договорами
'  21.3.12 - восстанавил запись OppN вместо OppId в Paid.csv
'  25.3.12 - диалог по связке Договора с Проектом
'   4.4.12 - поиск Организации, которой еще нет в SF
'  16.4.12 - создание новых Организаций в SF
'  12.5.12 - связь Платежей с Контрактами Autodesk; ревизия локальных Dim
'  20.5.12 - профилирование NewAcc/Add1CAcc

'---- декларация локальных переменных ----------
' в комментариях ниже "поле' - значение поля из листа Платежей1С
'                        =   - локальная переменная в этой процедуре
    Dim PayK As String      'поле - код Платежа
    Dim Dat As Date         'поле - "Дата прих.денег"
    Dim Acc As String       'поле - "Клиент"
    Dim Rub As Variant      'поле - "Итог руб"
    Dim Sale As String      'поле - "Продавец"
    Dim Good As String      'поле - "Товары" листа Платежей, спецификация
    Dim t As String         ' = Тип Товара по Спецификации
    Dim Sbs As Boolean      ' = True если Спецификация содержит Подписки
    Dim Dogovor As String   'поле - "Договор"
    Dim MainDog As String   'поле - "Осн.договор"
    Dim ContrK As String    ' = форма <Осн.договор>/<Договор>
    Dim ContrId As String   ' = Id Договора как в SF
    Dim OppId As String     ' = Id Проекта в SF
    
    Dim i                   ' = индекс основного цикла по Платежам
    Dim t0, Tbeg, TI    ' для профилирования по времени
    t0 = Timer
'-----------------------------------------------

    Dim ts1 As Long, ts2 As Long, ts3 As Long, ts4 As Long, ts5 As Long ' profiling ?????????
    Dim ts1S As Long, ts2S As Long, ts3S As Long, ts4S As Long, ts5S As Long ' profiling ?????????
    ts1 = 0: ts2 = 0: ts3 = 0: ts4 = 0: ts2 = 0: ts5 = 0

    Lines = ModStart(1, "PaidAnalitics: Просмотр Платежей", True) - 3
    
    ClearSheet O_NewOpp, Range("HDR_NewOpp")
    ClearSheet P_Paid, Range("HDR_Payment")
    ClearSheet C_ContrLnk, Range("HDR_ContrLnk")
    ClearSheet C_Contr, Range("HDR_NewContract")
    ClearSheet P_PaymentUpd, Range("HDR_PaymentUpd")
    ClearSheet A_Acc, Range("HDR_AdAcc")
    ClearSheet P_ADSKlink, Range("HDR_P_ADSKlink")
      
'    iniAdAccFr1C    ' инициализация хэш таблицы Организаций AdAcc
    
    TI = Timer
    LogWr t0 - TI & " PaidAnalitics: начало"
    Dim SumNewPay
    SumNewPay = 0
    
    For i = 2 To Lines
        Progress (i / Lines)
        If ExRespond = False Then GoTo Ex
        With Sheets(1)
        ' Account в SF есть, Платежа в SF нет, Док и Продавец не пустой
            Acc = Compressor(.Cells(i, PAYACC_COL)) ' Организация
            If .Cells(i, 1) = 1 And _
                    Trim(.Cells(i, PAYDOC_COL)) <> "" And _
                    Trim(.Cells(i, PAYSALE_COL)) <> "" Then
                Dat = .Cells(i, PAYDATE_COL)    ' дата Платежа
                Sale = .Cells(i, PAYSALE_COL)   ' Продавец
                Good = .Cells(i, PAYGOOD_COL)   ' Товар
                t = GoodType(Good)              ' Тип товара по Счету
                Sbs = IsSubscription(Good, t)   ' Подписка?
                Rub = .Cells(i, PAYRUB_COL)     ' сумма Платежа руб
                Dogovor = .Cells(i, PAYDOGOVOR_COL)
                MainDog = Mid(.Cells(i, PAYOSNDOGOVOR_COL), 9)
                ContrK = ContrCod(Dogovor, MainDog)
                ContrId = ContractId(ContrK)    ' Id Договора, если есть
                Dim IsInSF As String
                IsInSF = ""

    '================ Платеж еще не занесен в SF? =============================
                If .Cells(i, 4) <> 1 Then
                    
                    ts1 = Timer                     ' tttttttttttttttttttttttttttttttttttttt
                    
                    OppId = IsOpp(Sale, Acc, t, Rub, Dat, ContrK) ' Id Проекта в SF
                    If OppId = "" Then
                        NewOpp Acc, ContrK, Dat, Sale, Rub, "RUB", t, Sbs
                    Else
            '>>>>  занесение нового Платежа
                        NewPay i, OppNbyId(OppId), ContrId
            '>>>>  связь Проекта с Договором или создание нового Договора в SF
                        If ContrK <> "" And ContrId = "" Then
                            NewContract Dogovor, MainDog, ContrK
                        Else
                            ContrOppLink i, ContrK, ContrId, OppId
                        End If
                    End If
                    ts1S = ts1S + (Timer - ts1)      ' tttttttttttttttttttttttttttttttttttt
                    ts1 = ts1
                Else
    '================= Платеж есть в SF =======================================
                    ts2 = Timer                    ' tttttttttttttttttttttttttttttttttttt
                    PayK = .Cells(i, PAYCODE_COL)   ' код Платежа
                    IsInSF = " в SF "
                    If ContrId <> "" Then   ' Платеж в SF есть и он связан с Договором
            ' --- проверка, что в Платеже в SF есть правильная ссылка на Договор
                        If ContrNbyPay(PayK) <> ContrN(ContrK) Then
            '>>>>  связь Платежа с Договором
                                    ' нас не интересует, есть ли уже такая связка в SF -
                                    ' .. всегда связываем Платеж с договором по 1C
                            Call PaymentUpd(PayK, ContrId)
            '>>>> Договор Платежа связан с правильным Проектом? Если нет - пере-связываем
                            ContrOppLink i, ContrK, ContrId, OppIdbyPay(PayK)
                        End If
                    End If
            '>>>> Платеж связан с Контрактом Autodesk в SF? если нет - свяжем
                    If t = WE_GOODS_ADSK Then
                ' --- тут надо проверить есть ли уже связка !!!
                ' --- пока их вообще нет - проверяем PaidADSK всегда
                        PaidADSK PayK, Acc, Dat, Good, Sbs
                    End If
                    ts2S = ts2S + (Timer - ts2)      ' tttttttttttttttttttttttttttttttttttt
                    ts2 = ts2
                End If
'                If Not IsRightSale(Sale, T) Then
'                    LogWr "     WARNING:  Странный Платеж" & IsInSF & _
'                        ": строка в [1] = " & i & " -- Продавец = " & Sale _
'                        & ", не работает с Товаром = " & T
'                End If
            Else
'                ts3 = Timer                    ' tttttttttttttttttttttttttttttttttttt
'                If NewAcc(Acc) = "*" Then
                ts4 = Timer                    ' tttttttttttttttttttttttttttttttttttt
                AdAccFr1C Acc               ' такой Организации нет в SF - занесем!
                ts4S = ts4S + (Timer - ts4)      ' tttttttttttttttttttttttttttttttttttt
                ts4 = ts4
'                End If
'                ts3S = ts3S + (Timer - ts3)      ' tttttttttttttttttttttttttttttttttttt
'                ts3 = ts3
'If i = 100 Then Exit For
            End If
        End With
    Next i
    
'    MsgBox "SumNewPay = " & SumNewPay
    
    Dim ResultMsg
Ex: ResultMsg = "ИТОГ: занесено " & EOL_NewPay - 1 & " новых Платежей; " _
        & EOL_NewOpp - 1 & " новых Проектов; " _
        & EOL_PaymentUpd - 1 & " Платежей связано с Договорами; " _
        & EOL_ContrLnk - 1 & " Договоров связано с Проектами; " _
        & EOL_AdAcc - 1 & " новых Организаций;" _
        & EOL_ADSKlnkPay - 1 & " Платежей связано с Контрактами Autodesk;"
    LogWr ResultMsg
    MsgBox ResultMsg
 
    MsgBox "time1: " & ts1S & "; time2: " & ts2S & "; time3: " & ts3S & "; time4: " & ts4S
      
    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\Payment\"
    WriteCSV P_Paid, "Paid.txt"
    Shell "quota2.bat"
    WriteCSV P_PaymentUpd, "PmntUpd.txt"
    Shell "quota3.bat"
    
    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\OppInsert\"
    WriteCSV O_NewOpp, "OppInsert.txt"
    Shell "quota2.bat"
    
    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\Dogovor\"
    WriteCSV C_ContrLnk, "ContrUpd.txt"
    Shell "quota3.bat ContrUpd.txt C:\SFconstr\Dogovor\ContrUpd.csv"
    
    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\Account\"
    WriteCSV A_Acc, "AdAcc.txt"
    Shell "quota_Acc.bat"

    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\ADSK\P_ADSK"
    WriteCSV P_ADSKlink, "P_ADSKlink.txt"
    Shell "quota_P_ADSK.bat"

    ModEnd 1
End Sub
Sub NewPay(i, OppN, ContrId)
'
' новый Платеж в SF из строки i 1С - заносим посредством DL
'   21.2.12
'   28.2.12 - исп EOL_NewPay

    Dim j  ' последняя строка листа P_Paid
    
    j = EOL_NewPay + 1
    EOL_NewPay = j
    
    With Sheets(P_Paid)
        .Cells(j, 1) = Sheets(1).Cells(i, 6)            ' Плат.док.
        .Cells(j, 2) = DDMMYYYY(Sheets(1).Cells(i, 7))  ' Дата
        .Cells(j, 3) = Sheets(1).Cells(i, 8)            ' Счет
        .Cells(j, 4) = Dec(Sheets(1).Cells(i, 18))      ' Итог руб.
        .Cells(j, 5) = Sheets(1).Cells(i, 19)           ' Товар
        .Cells(j, 6) = ContrId                          ' ContractId
        .Cells(j, 7) = OppN                             ' OppN
    End With
End Sub
Sub NewOpp(Account, ContrK, CloseDate, Sale, value, CurrencyOpp, TypGood, Sbs, _
    Optional Stage = "90%-первые деньги пришли на счет")
'
' новый проект для записи DL в Организации Account.
'   12.2.2012
'   12.3.12 - заполнение поля Компаньон = OppBuddy
'   13.3.12 - минимальный размер NewOpp
'   23.3.12 - возможность остановки через Public ExRespond=False
'   22.4.12 - bug fix с Dec(Value)
'   24.4.12 - логика работы с Line, Kind, OppType

    Dim i, AccountId, OwId, OwnerOpp, N, V
    Dim OppBuddy, OppName, OppUniq As String
        
    AccountId = AccId(Account)
    OwId = OwnerId(Sale, OppBuddy)
    OppUniq = Account & "-" & TypGood & " " & ContrK
    OppName = OppUniq
'---- диалог при создании нового проекта ----------------
    If TypGood <> "Расходники" Then
        OppName = OppUniq & " " & CloseDate
        If value < MinNewOpp Then Exit Sub
        If value < MinNewOppDialog Then
            Dim R
            R = MsgBox("Проект '" & OppName & "'" & vbCrLf & vbCrLf _
                & value & " " & CurrencyOpp _
                & " -- Создаем? ", vbYesNoCancel)
            If R <> vbYes Then
                If R = vbCancel Then ExRespond = False
                Exit Sub
            End If
        End If
    Else
'---- дедупликация Проектов по Расходникам:
'           В Организации отдаленная дата разрешена только для Расходников
        With Sheets(SFopp)
            For i = 1 To EOL_SFopp
                If .Cells(i, SFOPP_ACC1C_COL) = Account _
                        And .Cells(i, SFOPP_CLOSEDATE_COL) >= DATE_BULKY _
                    Then Exit Sub
            Next i
        End With
    End If
    
    With Sheets(O_NewOpp)
'---- дедупликация вновь создаваемых Проектов -----------
        For i = 1 To EOL_NewOpp
            If .Cells(i, 1) = AccountId Then    ' если мы уже создали подобный NewOpp
            
                If .Cells(i, NEWOPP_OPPUNIQ_COL) = OppUniq Then
                    If TypGood = "Расходники" Then
                        Exit Sub
                    Else
                        V = WorksheetFunction.Substitute(.Cells(i, NEWOPP_VALUE_COL), ".", ",")
                        .Cells(i, NEWOPP_VALUE_COL) = Dec(value + V)
                    End If
                    Exit For
                End If
            End If
        Next i

' ----- определяем Направление деят., Тему и Тип Проекта по TypGood ----
        Dim Line, Kind, OppType As String
        Dim iG As Range
        For Each iG In Range("GoodSbs").Rows
            If iG.Cells(1, WE_GOODS_COL) = TypGood Then
                Line = iG.Cells(1, WE_GOODS_LINCOL)
                Kind = iG.Cells(1, WE_GOODS_KINDCOL)
                OppType = TypGood                       ' инициализация типа Проекта
                If TypGood = "Оборудование" Then OppType = "Железо"
                If iG.Cells(1, WE_GOODS_ISSBSCOL) <> "" Then
                    OppType = iG.Cells(1, WE_GOODS_ISSBSCOL)
                    If Not Sbs Then OppType = iG.Cells(1, WE_GOODS_NOSBSCOL)
                End If
                Exit For
            End If
        Next iG
       
' ----- вставка нового Проекта в NewOpp, заполонение полей -------------
        EOL_NewOpp = EOL_NewOpp + 1
        N = EOL_NewOpp
        .Cells(N, NEWOPP_ACCID_COL) = AccountId
        .Cells(N, NEWOPP_OPPNAME_COL) = OppName
        .Cells(N, NEWOPP_CLOSDATE_COL) = DDMMYYYY(CloseDate)
        .Cells(N, NEWOPP_OWNERID_COL) = OwId
        .Cells(N, NEWOPP_VALUE_COL) = Dec(value)
        .Cells(N, NEWOPP_ISOCUR_COL) = CurrencyOpp
        .Cells(N, NEWOPP_TYPOPP_COL) = "Лицензии"
        .Cells(N, NEWOPP_STAGENAME_COL) = Stage
        .Cells(N, NEWOPP_TYPOPP_COL) = OppType
        .Cells(N, NEWOPP_LINE_COL) = Line
        .Cells(N, NEWOPP_KIND_COL) = Kind
        .Cells(N, NEWOPP_BUDDY_COL) = OppBuddy          ' Компаньон
        .Cells(N, NEWOPP_OPPUNIQ_COL) = OppUniq         ' служебное поле OppUniq
        If TypGood = "Расходники" Then
            .Cells(N, NEWOPP_CLOSDATE_COL) = "1.1.2020"
            .Cells(N, NEWOPP_VALUE_COL) = Dec(999999)
        End If
    End With
End Sub
Function IsOpp(Sale, Account, t, Rub, Dat, ContrCod)
'
' проверка, есть ли в Организации Account не полностью оплаченный проект типа Т.
' Если Проект не доплачен на сумму, меньшую Rub - возвращает Id этого проекта.
' Эта же подпрограмма находит подходящий Проект для связи его с Договором,
' указанном в TargetContrK по контекстному поиску имени Договора в названии Проекта.
' Если нет - возвращает ""
'   13.2.2012
'   21.2.2012 - доплата по закрытым Проектам
'   28.12.12 - выбор подходящего Проекта для внесения Платежа
'   3.3.12 - обработка даты создания Проекта OppCrDate,
'   5.3.12 - Optional ContrCod для проверки наличия более общего Проекта в SF
'   9.3.12 - не ищем Проект, если профиль Продавца не включает Товар Т
'  16.3.12 - обходим Проекты Closed/Lost, то есть Вероятность = 0%
'  18.3.12 - возвращаем Id а не номер найденного Проекта
'  19.3.12 - TargetContrK - целевой Договор для связи с Проектом Opp
'  22.3.12 - сверяем тип Т с типом Проекта - если разные - это неподходящий Opp
'  31.3.12 - переписана логика выбора подходящего Проекта
'  20.4.12 - bug fix для Платежей с Договором
'  25.4.12 - Платеж с Договором обрабатывается в отдельном проходе по Проектам

    Dim i, OppToPayRub, OppCur, OppN, ContrOpp, OppName, OppId
    Dim OppCloseDate As Date, OppCrDat As Date
    
    IsOpp = ""
    
    If t = "" Then Exit Function
    If Not IsRightSale(Sale, t) Then Exit Function

    Dim SeekOppType As String, OppT As String
    SeekOppType = WorksheetFunction.VLookup(t, Range("OppTypeRng"), 4, False)
 
    With Sheets(SFopp)
'-- если Проект связан с Договором -- проверим, что связь с ContrCod есть
        If ContrCod <> "" Then
            For i = 2 To EOL_SFopp
                If .Cells(i, 4) = Account Then
                    OppName = .Cells(i, SFOPP_OPPNAME_COL)          ' имя Договора
                    OppId = .Cells(i, SFOPP_OPPID_COL)
                    If InStr(OppName, ContrCod) <> 0 Then GoTo Found    '  есть в Проекте?
                    If IsRightContrOppLink(OppId, ContrCod) Then GoTo Found
                End If
            Next i
    ' -- сюда попадаем, если все Проекты SF просмотрели, но нужного Договора не нашли
    '    в этом случае пытаемся подобрать Проект подходящий для связи с Договором
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
                    And .Cells(i, SFOPP_PROBABILITY_COL) > 0 _
                    And IsSameTeam(Sale, .Cells(i, SFOPP_SALE_COL), OppN) _
                    And OppToPayRub >= Rub _
                    And Dat <= OppCloseDate Then GoTo Found
            End If
        Next i
    End With
    Exit Function
Found:
    IsOpp = Sheets(SFopp).Cells(i, SFOPP_OPPID_COL)    ' Есть Проект нужного типа
    
    Const ErMsg = " <!> WARNING IsOpp: необходима ревизия Проекта Id= "
    If OppToPayRub < Rub Then _
        LogWr ErMsg & OppId & "(" & OppN & ") Платеж " & Rub & " превышает размер Проекта"
    If OppCloseDate < Dat Then _
        LogWr ErMsg & OppId & "(" & OppN & ") Дата " _
            & DDMMYYYY(Dat) & " позже Даты закрытия Проекта " & DDMMYYYY(OppCloseDate)
    If InStr(OppT, SeekOppType) = 0 Then _
        LogWr ErMsg & OppId & "(" & OppN & ") тип Платежа '" & t _
            & "' не соответствует типу Проекта '" & OppT & "'"
End Function
Sub ContrOppLink(iPay, ContrK, ContrId, OppId)
'
' создание связи Договор - Проект по Платежу в строке iPay
'   25.3.12
'   25.4.12 - replace Vlookup на цикл

    Dim Contr_Opp, Contr_Pay
    Dim PayCode, OpNm, Resp, MSG As String
    
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
    MSG = "Договор '" & ContrK & "'"
    If Contr_Opp = 0 Then
        MSG = MSG & " не связан с Проектом."
    Else
        MSG = MSG & "связан с Проектом '" _
            & OppNameByN(Contr_Opp) & "' (" & Contr_Opp & ")"
    End If
    MSG = MSG & vbCrLf & vbCrLf _
        & "Предлагается связь '" _
        & OpNm & "' (" & Contr_Pay & ")"
    If Contr_Opp = 0 Then
        MSG = MSG & vbCrLf & vbCrLf & "Свяжем ?"
    Else
        MSG = MSG & vbCrLf & vbCrLf & "Заменим ?"
    End If
    Resp = MsgBox(MSG, vbYesNoCancel)
    If Resp = vbCancel Then ExRespond = False
    If Resp = vbYes Then
        EOL_ContrLnk = EOL_ContrLnk + 1
        With Sheets(C_ContrLnk)
            .Cells(EOL_ContrLnk, 1) = ContrId
            .Cells(EOL_ContrLnk, 2) = OppId
        End With
    End If
End Sub

Sub P_PaidContract()
'
' обработка новых Платежей по Договорам, формирование файла ввода для Payment
'   5.1.2012
'   10.1.2012 - формирование P_PaidContract по листу Платежей
'   14.2.2012 - вместо формул макро с проходом по Платежам
    
    Const WSheetName = "P_PaidContract" ' имя рабочего листа в Excel
    
    ModStart WSheetName, "Новые Платежи по Договорам в файле " & WSheetName
    
' готовим файл Платежей
    AutoFilterReset 1
    
    Set Payments = ActiveSheet.Range("A1:AC" & ActiveSheet.UsedRange.Rows.Count)
    Payments.AutoFilter Field:=25, Criteria1:="<>"  ' Платежи 1С, где указан Договор
    Payments.AutoFilter Field:=4, Criteria1:="="    '   .. которых еще нет в SF
    Payments.AutoFilter Field:=1, Criteria1:="<>"   '       .. а Организация есть
       
    Range("B1:B" & ActiveSheet.UsedRange.Rows.Count - 3).Copy ' копируем выбранные индексы
    
    Sheets(WSheetName).Activate
    Range("A3").Select
    ActiveSheet.Paste       ' копируем выбранные Платежи после 2й строки P_PaidContract
    
    Lines = Sheets(WSheetName).UsedRange.Rows.Count - 6
    Range("C2:P" & Lines + 2).Select
    Selection.FillDown                  ' дополняем колонки формул
    Selection.RowHeight = 15
    
    Rows("2:3").Delete                  ' удаляем строки старых данных и заголовок из Платежей
    Range("A2:A" & Lines).Interior.Color = Gray
    Rows(Lines + 1 & ":1000").Delete
    
' ????? почему-то не работает фильтр - проблемы с форматом OppId в Excel - фильтрую вручную
'    Set PaidContr = ActiveSheet.Range("A1:I" & ActiveSheet.UsedRange.Rows.Count)
'    PaidContr.AutoFilter Field:=9, Criteria1:="<>0"       ' Платежи 1С, где указан Договор

    For i = 2 To Lines
        If Range("I" & i).value = 0 Then Rows(i).Hidden = True
    Next i
    
    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\Payment"
    WritePaid "Paid.txt"
    Shell "quota2.bat"
    
    AutoFilterReset 1
    ModEnd WSheetName
End Sub
Function GoodType(G) As String
'
' возвращает тип товара G по таблице в We.
' если подходящий тип не найден - ошибка и GoodType = ""
'   19.2.2012

    Dim j As Integer
    Dim iG As Range
    Dim S, Goods() As String
    
    GoodType = ""
    If G = "" Then Exit Function
    For Each iG In Range("Goods").Rows
        GoodType = iG.Cells(1, 1)
        S = iG.Cells(1, 2)
        Goods = split(S, ",")           ' в Goods список товаров данного типа
'If GoodType = "О П Л А Т А" Then
'j = j
'End If
        For j = 0 To UBound(Goods)
            If InStr(G, Trim(Goods(j))) > 0 Then Exit Function
        Next j
    Next iG
    LogWr "ВНИМАНИЕ! Неизвестный тип товара " & G
End Function
Function IsSubscription(Good, GT) As Boolean
'
' возвращает True, если товар - подписка/Subscription/Maintanence
' в зависимости от типа товара GT. Иначе - поставка лицензии, т.е. False.
' Определяется по таблице товаров в We
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
    LGood = LCase$(Good)
    
    SbsWords = split(LCase$(Sbs), ",")
    For i = LBound(SbsWords) To UBound(SbsWords)
        If InStr(LGood, Trim(SbsWords(i))) > 0 Then
            IsSubscription = True
            Exit Function
        End If
    Next i
    
End Function
