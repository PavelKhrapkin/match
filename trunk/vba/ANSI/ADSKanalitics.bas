Attribute VB_Name = "ADSKanalitics"
'---------------------------------------------------------------------------------
' ADSKanalitics  - анализ спецификаций продуктов Autodesk в Счете 1С---
'
' - PaidADSK(PayK, Acc, Dat, Spec,Sbs) - анализ спецификации Spec, ее связь с Autodesk
' - ContrADSKbySN(SN)           - возвращает Контракт Autodesk по SN в SF
' - ADSKqty(Acc, Desck, Dat, Contr) - сколько у Организации Acc мест типа Descr
' T testDIC_GoodADSK()          - отладка IsGoodInSpec и FindInLst
' - IsGoodInSpec(Good, Spec)    - распознает есть ли товар Good в Spec
' - IsContrADSKinSF(ContrADSK)  - возвращает TRUE если контракт ContrADSK есть в SF
'
'   9.6.2012

Option Explicit

Sub PaidADSK(PayK, Acc, Dat, Spec, Sbs)
'
' - PaidADSK(PayK, Acc,Dat,Spec,Sbs)  - анализ спецификации Spec, ее связь с Autodesk
'   14.5.12
'    9.6.12 - ревизия

'------ проверим проводку товара ADSK по складской книге и по ADSKfrSF
    Dim SNstock As String       '= SN проведенный по Складу для этого Платежа
    Dim StockDeskr As String    '= Продукт ADSK по этому SN
    Dim StockRec As String      '= возвращаемое поле Клиент по Складу
    
    SNstock = SN_ADSKbyStock(PayK, Acc, Dat, StockRec)
    If SNstock = "" Then Exit Sub
        
    Dim i As Integer
    Dim j As Integer
    Dim L As Integer
    Dim R As String
        
    Dim SpecLine() As String
    Dim Descr As String
    Dim Qty As Integer
    Dim PayId As String
    Dim ContrADSK As String, ContrId As String
    Dim SNarray() As String

        
    PayId = PayIdByK(PayK)
    
    SpecLine = split(Spec, ";") ' строки спецификации отделяются ;
'----- выделим тип продукта ADSK и количество в строке спецификации
    For i = LBound(SpecLine) To UBound(SpecLine)
        R = SpecLine(i)
        If R = "" Then Exit For
        L = Len(R)
        Descr = FindInLst(SpecLine(i), "DIC_GoodADSK")
        For j = 1 To 5              ' не более 5 цифр - шт. для позиции Счета
            If Mid(R, L - j, 1) = "/" Then
                Qty = Right(R, j)
                Exit For
            End If
        Next j

        
' SNstock может быть строкой вида "123-4356789+123-456789" -- разбираем по одному
        SNarray = split(SNstock, "+")
        For j = LBound(SNarray) To UBound(SNarray)
            StockDeskr = LCase$(FindInLst(ProdADSKbySN(SNarray(j)), "DIC_GoodADSK"))
            If InStr(LCase$(Spec), StockDeskr) <> 0 Then
                ContrADSK = ContrADSKbySN(SNarray(j), ContrId)
'                If ADSKidByPayId(PayId) = ContrId Then Exit Sub   ' если связь уже есть - уходим
                If IsADSK_PA(ContrADSK, PayId) Then Exit Sub   ' если связь уже есть - уходим
                Call UpdLinkADSK(StockRec, PayId, ContrId)
            Else
'                MsgBox "По Складу SN=" & SNstock & " (" & StockDeskr & "), " _
'                    & vbCrLf & vbCrLf & " а в Счете '" & Spec & "'" _
'                    & vbCrLf & vbCrLf & "нет продуктов ADSK этой группы!"
'                Stop
            End If
        Next j
    Next i
'    If ADSKidByPayId(PayId) <> "" Then Exit Sub   ' если связь уже есть - уходим
            
''------ есть подходящий продукт ADSK вида Descr с количеством Qty в SF?
'        If Descr <> "" Then
'            Dim Resp As String
'            Dim SN As String
'            Dim Contr As String
'
'            If Qty = ADSKqty(Acc, Descr, Dat, Contr, ContrId) Then
'                Resp = MsgBox("В '" & Acc & "' есть " & Qty & vbTab & Descr _
'                    & vbTab & Contr & vbCrLf & vbCrLf & R _
'                    & vbCrLf & vbCrLf & "Свяжем?", vbYesNoCancel)
'                If Resp = vbCancel Then ExRespond = False
'                If Resp = vbYes Then Call UpdLinkADSK(PayId, ContrId)
'            End If
'        End If
End Sub
Function ContrADSKbySN(SN, ContrId) As String
'
' - ContrADSKbySN(SN) - возвращает Контракт Autodesk по SN в SF
'   28/5/12

    Dim i As Integer
    
    ContrADSKbySN = ""
    With Sheets(ADSKfrSF)
        For i = 2 To EOL_ADSKfrSF
            If SN = .Cells(i, SFADSK_SN_COL) Then
               ContrADSKbySN = .Cells(i, SFADSK_CONTRACT_COL)
               ContrId = .Cells(i, SFADSK_CONTRID_COL)
               Exit Function
            End If
        Next i
    End With
    LogWr "ADSK FATAL ERR: SN#" & SN & " не найден"
End Function
Function ProdADSKbySN(SN) As String
'
' - ProdADSKbySN(SN) - возвращает Product Description продукта Autodesk по SN.
'   13/5/12

    ProdADSKbySN = ""
    On Error Resume Next
    ProdADSKbySN = WorksheetFunction.VLookup(SN, _
        Sheets(ADSKfrSF).Range("D:E"), 2, False)
    On Error GoTo 0
End Function
Function IdContrADSK(ContrADSK) As String
'
' - IdContrADSK(ContrADSK) - возвращает Id SF Контракта Autodesk.
'   23/5/12
'   8.6.12 - bug fix

    Dim Ctr() As String
    Dim i As Integer

    IdContrADSK = ""
    Ctr = split(ContrADSK, "+")
    For i = LBound(Ctr) To UBound(Ctr)
        Ctr(i) = Replace(Ctr(i), "'", "")
'        If Left(Ctr(i), 1) = "'" Then Ctr(i) = Mid(Ctr(i), 1, 12)
        On Error Resume Next
        IdContrADSK = WorksheetFunction.VLookup(Ctr(i), _
            Sheets(ADSKfrSF).Range("A:B"), 2, False)
        On Error GoTo 0
        If IdContrADSK <> "" Then Exit Function
    Next i
End Function
Function IsADSK_PA(ContrADSK, PayId) As Boolean
'
' - IsADSK_PA(ContrADSK, ) - возвращает TRUE если ContrADSK связан с Платежом
'   23.5.12

    Dim i As Integer
    Dim ContrId As String
    
    IsADSK_PA = False
    
    ContrId = IdContrADSK(ContrADSK)
    
    With Sheets(SF_PA)
        For i = 2 To EOL_SFlnkADSK
            If ContrId = .Cells(i, SFPA_ADSKID_COL) _
                    And PayId = .Cells(i, SFPA_PAYID_COL) Then
                IsADSK_PA = True
                Exit Function
            End If
        Next i
    End With
End Function
Function ADSKidByPayId(PayId) As String
'
' ADSKidByPayId(PayId) - возвращает Id SF Контракта Autodesk
'        или "", если Платеж еще не связан с Контрактом ADSK
'   15/5/12

    Dim i As Integer
    
    ADSKidByPayId = ""
    With Sheets(SF_PA)
        For i = 2 To EOL_SFlnkADSK
            If PayId = .Cells(i, SFPA_PAYID_COL) Then
               ADSKidByPayId = .Cells(i, SFPA_ADSKID_COL)
               Exit For
            End If
        Next i
    End With
End Function
Sub UpdLinkADSK(StockRec, PayId, ADSKcontrId)
'
' - UpdLinkADSK(PayId, ADSKcontrId) - занесение новой связи в лист P_ADSKlink
'   14.5.12
'   21.5.12 - выход, если PayId или ADSKcontrId пустые
   
' ---- дедупликация связок Платежей с Контрактом ADSK
    If PayId = "" Or ADSKcontrId = "" Then Exit Sub
    Dim i As Integer
    With Sheets(P_ADSKlink)
        For i = 2 To EOL_ADSKlnkPay
            If .Cells(i, 1) = StockRec _
                    And .Cells(i, 2) = PayId _
                    And .Cells(i, 3) = ADSKcontrId _
            Then Exit Sub
        Next i
    End With
    
    EOL_ADSKlnkPay = EOL_ADSKlnkPay + 1
    
    With Sheets(P_ADSKlink)
        .Cells(EOL_ADSKlnkPay, 1) = StockRec
        .Cells(EOL_ADSKlnkPay, 2) = PayId
        .Cells(EOL_ADSKlnkPay, 3) = ADSKcontrId
    End With

End Sub
Function ADSKqty(Acc, Descr, Dat, Contr, ContrId) As Integer
'
' ADSKqty(Acc, Desck, Dat, Contr, ContrId) - сколько у Организации Acc мест типа Descr
'                           Contr - возвращаемый номер Контракта Autodesk и его Id
'                           Учитываем только Registered SN с подпиской.
'   8.5.12

'-----------------------
' по видимому стоит завести в SF объект для связи SN с Платежом
' это позволит анализировать не только факт продления контракта ADSK,
' но и проверять продление отдельных SN
'-----------------------
' пока не разбираем к каким Контрактам ADSK относятся продукты типа Desct
' и не следим за Dat - просто суммируем. Потом надо изощренее.
'----------------------------------------------
    
    Dim i As Integer
    Dim SN As String
    Dim ProdSN As String
    Dim ContrSN As String

    ADSKqty = 0
    Contr = ""
    With Sheets(ADSKfrSF)
        For i = 2 To EOL_ADSKfrSF
            If Acc = .Cells(i, SFADSK_ACC1C_COL) Then
                SN = .Cells(i, SFADSK_SN_COL)
                
                If .Cells(i, SFADSK_STATUS_SN_COL) = SFADSK_SN_REGISTERED Then
                    ProdSN = .Cells(i, SFADSK_DESCRIPTION_COL)
                    If Descr = FindInLst(ProdSN, "DIC_GoodADSK") Then
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' дисциплина работы с датами по продуктам Autodesk еще не продумана!!!!
                        If Dat >= .Cells(i, SFADSK_CONTRSTARTDAT_COL) Then
                            ADSKqty = ADSKqty + .Cells(i, SFADSK_SEATS_COL)
                            ContrSN = .Cells(i, SFADSK_CONTRACT_COL)
                            If Contr = "" Then
                                Contr = ContrSN
                                ContrId = .Cells(i, SFADSK_CONTRID_COL)
                            ElseIf Contr <> ContrSN Then
'                              MsgBox "Организация '" & Acc & "': " _
'                                    & Descr & " в разных Контрактах Autodesk!"
                                Contr = "": ContrId = "": ADSKqty = 0
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    End With
End Function
Sub testDIC_GoodADSK()
'
' T testDIC_GoodADSK() отладка IsGoodInSpec и FindInLst
'   10/5/12

    Const Spec = "Программное обеспечение AutoCAD Architecture Commercial Subscription (1 year) (Renewal)/3;Программное обеспечение AutoCAD Inventor Professional Suite 2011 Upgrade from Inventor Suite 2011 RU/2;Программное обеспечение Autodesk Inventor Professional Commercial Subscription (1 year)/2;Программное обеспечение AutoCAD Inventor Suite Subscription Renewal/3;Программное обеспечение AutoCAD MEP Commercial Subscription (1 year)/4;Программное обеспечение AutoCAD Commercial Subscription (1 year) (Renewal)/1;Программное обеспечение AutoCAD Civil 3D Commercial Subscription  (Renewal)/2;Программное обеспечение AutoCAD Electrical Commercial Subscription (1 year) (Renewal)/6;"
    Const good = "AutoCAD Revit Architecture Suite 2012 Russian"
    
    Dim A, i, GoodADSK, R
    Dim CSD_Line As Range
    
    A = FindInLst("Autodesk Product Design Suite Premium 2012 Commercial New SLM", "DIC_GoodADSK")
    
    A = IsGoodInSpec(good, Spec)
    
    ModStart "We", "Проверка DIC_GoodADSK по списку из CSD"
    For Each CSD_Line In Range("DIC_Build_Autodesk_Material_Description").Rows
        GoodADSK = CSD_Line.Cells(1, 2)
        A = FindInLst(GoodADSK, "DIC_GoodADSK")
'        R = MsgBox(A & vbTab & GoodADSK, vbYesNo)
        CSD_Line.Cells(1, 1) = A
'        If R = vbYes Then
'            CSD_Line.Cells(1, 1) = A
'        Else
'            CSD_Line.Cells(1, 1) = "No!"
'         End If
    Next CSD_Line
''''''''''''''''''''''''''''''''
    R = MsgBox("Проход по Платежам - выполним?", vbYesNo)
    If R = vbNo Then GoTo ADSFfrSF_Pass

    Lines = ModStart(PAY_SHEET, "Проход по Платежам", True)
Lines = 100
    For i = 2 To Lines
        Progress i / Lines
        GoodADSK = Sheets(PAY_SHEET).Cells(i, PAYGOOD_COL)
        If GoodType(GoodADSK) = WE_GOODS_ADSK Then
            A = FindInLst(GoodADSK, "DIC_GoodADSK")
            If A <> "" Then
    '            MsgBox A & vbTab & GoodADSK
            Else
                MsgBox "Не распознан '" & GoodADSK & "'"
            End If
        End If
    Next i
    ModEnd PAY_SHEET
'''''''''''''''''''''''''''''''
ADSFfrSF_Pass:

    R = MsgBox("Проход по ADSKfrSF - выполним?")
    If R = vbNo Then Stop
    
    ModStart ADSKfrSF, "DIC_ADSK test", True
EOL_ADSKfrSF = 100
    For i = 2 To EOL_ADSKfrSF
        GoodADSK = Sheets(ADSKfrSF).Cells(i, SFADSK_DESCRIPTION_COL)
        A = FindInLst(GoodADSK, "DIC_GoodADSK")
        If A <> "" Then
'            MsgBox A & vbTab & GoodADSK
        Else
            MsgBox "Не распознан '" & GoodADSK & "'"
        End If
    Next i
    ModEnd ADSKfrSF
End Sub
Function IsGoodInSpec(good, Spec) As Boolean
'
' - IsGoodInSpec(Good, Spec)    - распознает есть ли товар Good в Spec
'                  распознование товара происходит по словарю DIC_GoodADSK
'   7.5.12

    Dim Gkey As String
    Dim Skey As String
    Dim SpecLines() As String
    Dim i As Integer
    
    IsGoodInSpec = False
    If good = "" Or Spec = "" Then Exit Function
    
    Gkey = FindInLst(good, "DIC_GoodADSK")
    
    SpecLines = split(Spec, ";")
    For i = LBound(SpecLines) To UBound(SpecLines)
        Skey = FindInLst(SpecLines(i), "DIC_GoodADSK")
        If Gkey = Skey Then
            IsGoodInSpec = True
            Exit Function
        End If
    Next i
End Function
Sub PutInTab(Rng, Val, Cat, Dat)
'
'- PutInTab(Rng, Val, Cat, Dat) - заносит в Range Rng Val в поле по Cat и Dat
'   15/5/12

    Dim i As Integer
    Dim Str As Range
    Dim h As Range
    Dim HDR_Date As Date
    Dim y, Y0, M, M0
    
    If Not IsDate(Dat) Then GoTo Error
    Y0 = Year(Dat)
    M0 = Month(Dat)
       
    For Each Str In Range(Rng).Rows
        If Str.Cells(1, 1) = Cat Then
            i = 1
            For Each h In Range(Rng).Columns
                If IsDate(h.Cells(1, 1)) Then
                    y = Year(h.Cells(1, 1))
                    M = Month(h.Cells(1, 1))
                    If y = Y0 And M = M0 Then
                        Str.Cells(1, i) = Str.Cells(1, i) + Val
                        Exit Sub
                    End If
                End If
                i = i + 1
            Next h
        End If
    Next Str

Error: MsgBox Cat & " или " & Dat & " не найден!", , "ERROR"
    Stop
End Sub
Sub TestPutInTab()
'
' T TestPutInTab - получение статистики по Seats ADSK

'   15/5/12

'    Call PutInTab("ADSK_Lic", 32, "Plant", "15.10.09")
    
    Dim i, j, Qty As Integer, good As String, Descr As String
    Dim Sbs As Boolean, Consulting As Boolean
    
    Dim Dat As Date
    
    Call ModStart(PAY_SHEET, "Статистика по Платежам Autodesk", True)

    With Sheets(PAY_SHEET)
        For i = 2 To EOL_PaySheet
            Progress i / EOL_PaySheet
            If Trim$(.Cells(i, PAYDOC_COL)) <> "" Then
                good = .Cells(i, PAYGOOD_COL)
                Dat = .Cells(i, PAYDATE_COL)
                If GoodType(good) = WE_GOODS_ADSK Then
                    For j = 0 To 999
                        Descr = ADSK_SpecItem(good, j, Sbs, Consulting, Qty)
                        If Descr = "" Then Exit For
                        If Sbs Then
                            PutInTab "ADSK_Subs", Qty, Descr, Dat
                        Else
                            If Not Consulting Then PutInTab "ADSK_Lic", Qty, Descr, Dat
                        End If
                    Next j
                End If
            End If
        Next i
    End With
    
    ModEnd PAY_SHEET
End Sub
Function ADSK_SpecItem(Spec, Nstr, Sbs, Consulting, Optional Qty As Integer) As String
'
' - ADSK_SpecItem(Spec, Nstr, [Qty]) - разбор строки Nstr спецификации.
'               Возвращает тип этой строки ПО ADSK и количество Qty
'   30.5.12

    Dim SpecLine() As String
    Dim R As String
    Dim i As Integer
    Dim L As Integer
    
    ADSK_SpecItem = ""
    SpecLine = split(Spec, ";") ' строки спецификации отделяются ;
    
    If Nstr < LBound(SpecLine) Or Nstr > UBound(SpecLine) Then Exit Function
            
    R = SpecLine(Nstr)
    If R = "" Then Exit Function
    L = Len(R)
    ADSK_SpecItem = FindInLst(R, "DIC_GoodADSK")
    
    Sbs = False: Consulting = False
    
    If InStr(R, "Subscription") <> 0 Then Sbs = True
    If InStr(R, "Консультац") <> 0 Then Consulting = True
    
    For i = 1 To 5   ' не более 5 цифр - шт. для позиции Счета
        If Mid(R, L - i, 1) = "/" Then
            Qty = Right(R, i)
            Exit For
        End If
    Next i
End Function
Sub testADSK_SpecItem()

    Const Spec = "Программное обеспечение Autodesk Plant Design Suite Premium 2012, коммерческий, русский, сетевой/2;Программное обеспечение Autodesk Plant Design Suite Premium Subscription/2;Программное обеспечение Autodesk Plant Design Suite Ultimate 2012, коммерческий, русский, сетевой/1;Программное обеспечение Autodesk Plant Design Suite Ultimate Subscription/1;"
    
    Dim A(1 To 10) As String
    Dim Qty As Integer
    Dim Sbs As Boolean
    
    A(1) = ADSK_SpecItem(Spec, 2, Sbs)
    A(2) = ADSK_SpecItem(Spec, 1, Sbs, Qty)
End Sub
Function DateFrUSA(d) As String
'
' - DateFrUSE(D)   - преобразование текста D вида "M/D/YYYY" в дату "DD.MM.YY"
'   25.5.12

    Dim Dat() As String
    Dim Delimeter As String
    
    If InStr(d, ".") <> 0 Then
        Delimeter = "."
    ElseIf InStr(d, "/") Then
        Delimeter = "/"
    Else
        MsgBox "Нераспознаваемая дата в строке '" & d & "'", , "ERROR"
        Stop
    End If
    
    Dat = split(d, Delimeter)
    DateFrUSA = Dat(1) & "." & Dat(0) & "." & Dat(2)
End Function
Sub testDateFrUSA()
    Dim A, b, c, d
    A = DateFrUSA("1/26/2009")
    b = DateFrUSA("32/11/09")
    d = DateFrUSA("12.4.11")
End Sub
Sub ADSK_Contract_Handle()
'
' (*) ADSK_Contract_Handle()    - обработка листа ADSK_Contract из PartnerCenter
'   26.5.12

    Dim i
    
    Lines = ModStart(ADSK_C, "Проход по ADSK_Contracts")
    
    CheckSheet ADSK_C, 1, 2, ADSK_C_STAMP
    
'---- преобразование колонки дат "Contract End Date"
    With Sheets(ADSK_C)
        For i = 2 To Lines
            .Cells(i, ADSK_C_ENDDATE_COL) = DateFrUSA(.Cells(i, ADSK_C_ENDDATE_COL))
        Next i
    End With
    
    SheetSort ADSK_C, ADSK_C_CONTR_COL  ' сортировка по номеру контракта
    
'---- дедупликация по "Contract End Date" с выбором наиболее создней даты
    Dim D0 As Date      '= дата в предыдущей строке
    Dim D1 As Date      '= дата в текущей строке
    Dim Contr As String '= текущий номер Контракта ADSK
    Dim AccountN As String  'поле - номер клиента в ADSK_C
    Dim AccCSN As String    'поле - номер клиента ADSK (AccCSN) в SF
    
    With Sheets(ADSK_C)
        i = 2
        Do
            D1 = .Cells(i, ADSK_C_ENDDATE_COL)
            Contr = .Cells(i, ADSK_C_CONTR_COL)
            If Contr = .Cells(i - 1, ADSK_C_CONTR_COL) Then
                If D1 < D0 Then
                    Rows(i & ":" & i).Delete
                Else
                    Rows(i - 1 & ":" & i - 1).Delete
                    D0 = D1
                End If
                Lines = Lines - 1
            Else
                D0 = D1
                i = i + 1
            End If
        Loop While i <= Lines
    
'---- проверка того, что в SF есть Kонтракта ADSK с правильной датой окончания
        Dim Msg As String
    
        For i = 2 To Lines
            Contr = .Cells(i, ADSK_C_CONTR_COL)
            AccountN = .Cells(i, ADSK_C_ACCN_COL)
            D0 = .Cells(i, ADSK_C_ENDDATE_COL)
            D1 = ContrADSK_EndDate(Contr, AccCSN)
            Msg = "ADSK_C: Контракт ADSK #"
            If D0 <> D1 Then
                Msg = Msg & Contr & " кончается " & D0 & ", а в SF " & D1
                MsgBox Msg, , "WARNING"
                LogWr Msg
                Stop
            End If
            If AccountN <> AccCSN Then
                Msg = Msg & Contr & " принадлежит " & AccountN & ", а в SF " & AccCSN
                MsgBox Msg, , "WARNING"
                LogWr Msg
                Stop
            End If
        Next i
        
    End With
    
    ModEnd ADSK_C
End Sub
Function ContrADSK_EndDate(Contr, AccCSN) As Date
'
' - ContrADSK_EndDate(Contr, AccCSN)  - находит дату окончания контракта ADSK в SF и AccCSN
'   26.5.12
    
    On Error Resume Next
    ContrADSK_EndDate = WorksheetFunction.VLookup(Contr, _
        Sheets(ADSKfrSF).Range("A:L"), 12, False)
    AccCSN = WorksheetFunction.VLookup(Contr, _
        Sheets(ADSKfrSF).Range("A:H"), 8, False)
    On Error GoTo 0
End Function
