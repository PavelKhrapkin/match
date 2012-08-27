Attribute VB_Name = "StockAnalitics"
'---------------------------------------------------------------------------------
' StockAnalitics  - анализ базы по Складу
'
' S   GetInv1C(InvCol,PayN_Col, _       - находит номер строки в Платежах PayN со Счетом,
'         StrInvCol,DateCol,[Str2Inv])    найденным по строкам, содержащих Счет и по дате.
' [*] StockHandling()   - проход по листу "Склад", поиск продуктов Autodesk
'  -  FindAcc1C(Client, Acc1C) - поиск Счета 1С указанного в Client для Acc1C
'  -  SeekInv(Str) - выделение Счета в текстовой строке Str
'  -  SNhandl(Acc1C, PayN, StockSN) - обработка SN в складской книге
'  ?  RemIgnoredSN(S)   - удаление игнорируемых слов и знаков в S по SN
' (*) Sndedup() - дедупликация SN найденных по Складу - времянка!
'  -  SN_ADSKbyStock(PayId, Acc, Dat) - возвращает SN продукта ADSK по Складу

'   26.8.2012

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
    Dim Dat As Date         'поле "Дата" в складской книге
    Dim good, t As String   ' Товар (спецификация) и Тип товара
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
If i >= 7766 Then
i = i
End If
            Client = .Cells(i, STOCK_CLIENT_COL)
            Dat = TxDate(.Cells(i, STOCK_DATE_COL))
            If Client = .Cells(i - 1, STOCK_CLIENT_COL) Then
                SameClient = True
            Else
                SameClient = False
                Acc1C = FindAcc(Client)                 ' поиск Организации
                PayN = FindPayN(Client, Acc1C, Dat)     ' поиск Счета/Платежа
            End If
            .Cells(i, STOCK_ACC_COL) = Acc1C
            If PayN > 0 And PayN <= EOL_PaySheet Then
                .Cells(i, STOCK_INVOICE_COL) = _
                        Sheets(PAY_SHEET).Cells(PayN, PAYINVOICE_COL)
                good = Sheets(PAY_SHEET).Cells(PayN, PAYGOOD_COL)
                t = GoodType(good)              ' Тип товара по Счету
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
' FindAcc1C(Client) - поиск Платежа по Счету1С упомянутому в Client для Acc1C
'   7.5.12
'  23.5.12 -на входе в Acc1C может быть несколько Организаций соединенных "+"
'           FindPayN ищет Счет по всем и оставляет в Acc1C только найденую
'  17.6.12 - выделен поиск Сч- в строке в отдельную подпрограмму

    Dim invoice As String
    Dim StockInv As String      '= Счет, выделенный из поля Клиент по Складу
    Dim D As Date               'поле "Дата" по Складу
    Dim AccInv As String        '= Имя Организации, полученное по Счету
    Dim i1C As Integer          '= номер строки в Платеже по Счету
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
                ErrMsg WARNING, "Заказчик по " & invoice & " " _
                    & AccInv & ", а по Складу " & Acc1C
            End If
        End If
    End If
    
    If Acc1C = "" Or Acc1C = "*" Then Exit Function
    FindPayN = i1C
    
    
'----- поиск Счета в Платежах -------
'
'    For i = 2 To EOL_PaySheet
'        Acc = Sheets(PAY_SHEET).Cells(i, PAYACC_COL)
'        If InStr(Acc1C, Acc) <> 0 Then
'            If InStr(Invoice, StockInv) <> 0 Then
'                For j = 4 To 7 ' разбираем строку Счета вида "Сч-123 ..."
'                    If Mid(Invoice, j, 1) = " " Then
'                        Inv = Mid(Invoice, 4, j - 4)
'                        Exit For
'                    End If
'                Next j
'                Inv = "Сч-" & Inv
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
' - SeekInv(Str) - выделение Счета в текстовой строке Str
'   17.6.12

    Dim Sch As String
    Dim S As String
    Dim i As Integer

    Dim StWord() As String  '= массив слов в Str
    
    SeekInv = ""
    If Trim(Str) = "" Then Exit Function
    
    S = Replace(LCase(Str), "_", " ")
    S = Replace(LCase(S), ")", " ")
    S = Replace(LCase(S), "(", " ")
    S = Replace(LCase(S), """", " ")
    StWord = split(S, " ")
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
    Sch = Abs(Sch)
    If Sch = 0 Or Sch >= 1000 Then Exit Function
    SeekInv = "Сч-" & Sch
End Function
Sub testSeekInv()
    Dim A(1 To 6) As String
    A(1) = SeekInv("Заказ ЗАО ""ЛИК-94"" Сч-267 от 07.10.11 Кириллова ")
    A(2) = SeekInv("Заказ ЗАО ""ЛИК-94"" С-267 от 07.10.11 Кириллова ")
    A(3) = SeekInv("Заказ ЗАО ""ЛИК-94"" С -267 от 07.10.11 Кириллова ")
    A(4) = SeekInv("Заказ ЗАО ""ЛИК-94"" Сч- 267 от 07.10.11 Кириллова ")
    A(5) = SeekInv("Заказ ЗАО ""ЛИК-94"" Сч - 267 от 07.10.11 Кириллова ")  '!!! не распознался!!!
End Sub
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
    Dim t, Q, R
    t = RemIgnoredSN("456 765-67812345")
    Q = RemIgnoredSN("")
    R = RemIgnoredSN("456-5654323 текст 456-апр-567")
End Sub

Function RemIgnoredSN(Str) As String
'
' - RemIgnoredSN(S)   - удаление игнорируемых слов и знаков в S
'   7.5.12
'   8.6.12 - удаляем короткие (<12 знаков) номера
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
'   18.6.12 - use TxDate

    Const MaxDeliveryDays = 70

    Dim StockDat As Date    ' = дата проводки товара по Складу
    Dim StockSN As String   ' = SN товара Autodesk по Складу
    Dim StockSch As Integer ' = номер Счета по Складу
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
