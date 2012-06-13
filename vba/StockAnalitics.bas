Attribute VB_Name = "StockAnalitics"
'---------------------------------------------------------------------------------
' StockAnalitics  - анализ базы по Складу
'
' [*] StockHandling()   - проход по листу "Склад", поиск продуктов Autodesk
'  -  FindAcc1C(Client, Acc1C) - поиск Счета 1С указанного в Client для Acc1C
'  -  SNhandl(Acc1C, PayN, StockSN) - обработка SN в складской книге
'  ?  RemIgnoredSN(S)   - удаление игнорируемых слов и знаков в S по SN
' (*) Sndedup() - дедупликация SN найденных по Складу - времянка!
'  -  SN_ADSKbyStock(PayId, Acc, Dat) - возвращает SN продукта ADSK по Складу

'   10.6.2012

Option Explicit

Sub StockHandling()
'
' [*] StockHandling()   - - проход по листу "Склад", поиск продуктов Autodesk
'   7.5.12
'  13.5.12 - в колонку Счет Склада пишем Счет из Платежей, а не номер строки
'  13.5.12 - добавлена вычисляемая колонка SN_SF - SN найденный в SF
'  23.5.12 - ревизия
'   7.6.12 - для ускорения счета обходим FindAcc и FindPayN для одинаковых Client

    Dim i As Integer
    Dim PayN As Integer
    Dim PayId As String     ' = Id Платежа в SF
    Dim Client As String    ' запись в складской книге, Клиент, Счет, Заказ
    Dim SameClient As Boolean
    Dim Acc1C As String     ' имя Организации в справочнике 1С
    Dim Good, T As String   ' Товар (спецификация) и Тип товара
    Dim StockSN As String   ' Складская запись об SN
    Dim SNinSF As String    ' SN уже занесенный в SF
    Dim NewSN As String     ' SN, которого еще нет в SF
    Dim ContrADSK As String ' Контракт ADSK - из SF или по Складу
    
    Lines = ModStart(STOCK_SHEET, "Проход по Складу: SN Autodesk", True)
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
                Acc1C = FindAcc(Client)                 ' поиск Организации
                PayN = FindPayN(Client, Acc1C)          ' поиск платежа
            End If
            .Cells(i, STOCK_ACC_COL) = Acc1C
            If PayN > 0 And PayN <= EOL_PaySheet Then
                .Cells(i, STOCK_INVOICE_COL) = _
                        Sheets(PAY_SHEET).Cells(PayN, PAYINVOICE_COL)
                Good = Sheets(PAY_SHEET).Cells(PayN, PAYGOOD_COL)
                T = GoodType(Good)              ' Тип товара по Счету
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
' FindAcc1C(Client) - поиск Платежа по Счету1С упомянутому в Client для Acc1C
'   7.5.12
'  23.5.12 -на входе в Acc1C может быть несколько Организаций соединенных "+"
'           FindPayN ищет Счет по всем и оставляет в Acc1C только найденую

    FindPayN = 0
    If Acc1C = "" Or Acc1C = "*" Then Exit Function
    
'----- выделение Счета из записи в складской книге Client ----
    Dim Invoice As String
    Dim SeekInv As String
    Dim Sch As String
    Dim i As Integer

    Dim StWord() As String
    StWord = split(LCase(Client), " ")
    For i = LBound(StWord) To UBound(StWord)
        Sch = StWord(i)
        If Left(Sch, 1) = Chr(99) Or Left(Sch, 1) = "с" Then ' Ru или En "с"
            Do
                Sch = Mid(Sch, 2)
                If Sch = "" Then Exit Do
                If Left(Sch, 1) <> "-" And IsNumeric(Sch) Then GoTo FoundSeekInv
            Loop
            If i < UBound(StWord) Then
                Sch = StWord(i + 1) ' иногда Сч-123 имеет форму 'C 123'
                If IsNumeric(StWord(i + 1)) Then GoTo FoundSeekInv
            End If
        End If
    Next i
    Exit Function
FoundSeekInv:
    SeekInv = Sch
    
'----- поиск Счета в Платежах -------
    Dim Inv As String, Acc As String
    Dim j As Integer
    
    For i = 2 To EOL_PaySheet
        Acc = Sheets(PAY_SHEET).Cells(i, PAYACC_COL)
        If InStr(Acc1C, Acc) <> 0 Then
            Invoice = Sheets(PAY_SHEET).Cells(i, PAYINVOICE_COL)
            If InStr(Invoice, SeekInv) <> 0 Then
                For j = 4 To 7 ' разбираем строку Счета вида "Сч-123 ..."
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
' SNhandl(Acc1C, PayN, StockSN) - обработка SN в складской книге
'   7.5.12
'  13.5.12 - добавлено поле SNinSF - SN уже занесенный в SF
'  18.5.12 - несколько SN оформляются в виде SN+SN
'   7.6.12 - удаляем из SNhandl SN найденный в SF

    If PayN = 0 Or StockSN = "" Then Exit Function

    Dim i As Integer
    Dim j As Integer
    Dim SN As String    ' SN в отчете ADSKfrSF
    Dim S As String
    Dim CtrADSK As String
    Dim AccStock As String
    
    SNhandl = "": SNinSF = "": ContrADSK = ""
    S = RemIgnoredSN(StockSN)
    If Len(S) < 12 Then Exit Function   ' нет SN
    
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
                S = Replace(S, SN, "")  ' исключим найденный SN и Contract Autodesk
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
    R = RemIgnoredSN("456-5654323 текст 456-апр-567")
End Sub

Function RemIgnoredSN(Str) As String
'
' - RemIgnoredSN(S)   - удаление игнорируемых слов и знаков в S
'   7.5.12
'   8.6.12 - удаляем короткие (<12 знаков) номера
'  10.6.12 - bug fix -- Replace 1 раз!

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
' Sndedup() - дедупликация SN найденных по Складу - времянка!
'   7.5.12

    Call SheetDedup("SN", 1)
End Sub
Function SN_ADSKbyStock(PayK, Acc, Dat, StockRec) As String
'
' - SN_ADSKbyStock(PayId, Acc, Dat, StockRec) - возвращает SN продукта ADSK
'         по Складской Книге. В StockRec возвращается поле "Клиент" по Складу.
'         Делается проверка корректности параметров Платежа PayId и данных по Складу.
'         Если проверка не прошла - возвращает "".
'   24.5.12

    Const MaxDeliveryDays = 70

    Dim DatStr As String    'поле - дата проводки товара по Складу
    Dim StockDat As Date    ' = дата проводки товара по Складу
    Dim StockSN As String   ' = SN товара Autodesk по Складу
    Dim StockSch As Integer ' = номер Счета по Складу
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
' - GetContrADSKfrStock (StockSN) - извлекение Контракта ADSK из Складсклй книги
'   18.5.12

    Dim i As Long
    Dim S As String
    
    GetContrADSKfrStock = ""
    
'!!!!!!!!!!! ЗАГЛУШКА !!!!!!!!!!!!!!!!!!!!!!!!
    If InStr(StockSN, "110000") = 0 Then Exit Function
    
    For i = 1 To Len(StockSN) - 12
        S = Mid(StockSN, i, 12)
        If IsNumeric(S) And InStr(S, "110000") = 1 Then
            GetContrADSKfrStock = S
            Exit Function
        End If
    Next i
End Function
