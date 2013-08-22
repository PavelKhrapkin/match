Attribute VB_Name = "ContrAnalitics"
'-----------------------------------------------------------------------------
' ContrAnalitics    - анализ Контрактов и относящихся к ним Проектов
' [*] ContrPass     - проход по листу Договоров (3) и создание новых Проектов
'  -  OppT(V1C)     - возвращает тип Проекта по полю Поставщик в Договоре 1С
' [*] DogOppLink    - проход по SFD и поиск подходящих Проектов для связи
'  -  IsSameVendor(OppType, V1C, ContrCode)    - возвращает True, если Тема&Вид
'                           Проекта OppType соответствует Поставщику по Договору в 1С
'   3.2.2013

Option Explicit
Sub NewContr(NewContract As String)
'
' S NewContr()  - просмотр Договоров 1С для занесения в SF новых через DL
' 18.8.2012
'  3.9.12 - StepIn
' 31.10.10 - fix Call WrNewSheet
'  9.11.12 - вызов NewSheet из NewContr

    StepIn
    
    Dim Dog As TOCmatch
    Dim i As Long
    
    NewSheet NewContract
    
    Dog = GetRep(DOG_SHEET)

    With DB_1C.Sheets(DOG_SHEET)
        For i = 2 To Dog.EOL
            Progress i / Dog.EOL
            If .Cells(i, DOGIDSF_COL) = "" And .Cells(i, DOGISACC_COL) <> "" Then
                WrNewSheet NewContract, DOG_SHEET, i
            End If
        Next i
    End With
End Sub
Sub PaidContr(ByVal NewPayment As String)
'
' S PaidContr()  - Занесение Платежа с Договором, связанным с Проектом
' 21.11.12
' 30.12.12 - NewOpp для Договоров без связки с Проектом

    StepIn
    
    Dim Paid As TOCmatch, ContrK As String, IsErr As Boolean, OppId As String
    Dim i As Long
    
    Const F_rqst = "SFD/" & SFD_COD_COL & ":" & SFD_OPPID_COL
    
    NewSheet NewPayment
    
    Paid = GetRep(PAY_SHEET)

    With DB_1C.Sheets(PAY_SHEET)
        For i = 2 To Paid.EOL
            Progress i / Paid.EOL
            If .Cells(i, PAYINSF_COL) <> 1 And .Cells(i, PAYDOGOVOR_COL) <> "" Then
                ContrK = ContrCod(.Cells(i, PAYDOGOVOR_COL), .Cells(i, PAYOSNDOGOVOR_COL))
                OppId = FetchDoc(F_rqst, ContrK, IsErr)
                If OppId = "" Then
'---- NEWOPP              лист для CSV    файл      строка -- НЕ НАПИСАНО ЕЩЕ!!!
'                    WrNewSheet NEW_OPP, PAY_SHEET, i, "HDR_NewOppBy"
                End If
                If Not IsErr Then WrNewSheet NewPayment, PAY_SHEET, i, OppId
            End If
        Next i
    End With
End Sub
Sub ContrPass()
'
' Проход по отчету Договоров и обзор/создание соответствующих Проектов
'   4.4.12
    
    Dim Contr As Range
    Dim i, OppN, OppIs, OppNew, NoOpp As Integer
    Dim IsGenOpp As Integer
    Dim DogSFstat, ContrK, ContrId, OppTyp, t, Stage As String
    Dim Acc, Dat, Sale, DogValue, DogCur, OppStage As String
    Dim VendorSF, Vendor1C As String
    
    Lines = ModStart(DOG_SHEET, "Проход по Договорам: связанные Проекты", True) - DOGRES
    EOL_DogSheet = Lines
    EOL_SFD = EOL(SFD) - SFresLines
    EOL_SFopp = EOL(SFopp) - SFresLines
    EOL_SFacc = EOL(SFacc) - SFresLines
    
    CheckSheet DOG_SHEET, 1, 10, Stamp1Cdog1
    CheckSheet SFD, EOL_SFD + 2, 3, SFcontrRepName
    CheckSheet SFopp, EOL_SFopp + 2, 1, SFoppRepName
'---------- проход по Договорам ------------------------
    OppIs = 0: OppNew = 0: NoOpp = 0: Fruitful = 0
    ClearSheet O_NewOpp, Range("HDR_NewOpp")
    ClearSheet NewContract, Range("HDR_NewContract")
    ClearSheet NewContractLnk, Range("HDR_ContrLnk")
    
    For i = 2 To Lines
        Progress i / Lines
        If ExRespond = False Then Exit For
        With Sheets(DOG_SHEET)
            DogSFstat = .Cells(i, DOGSFSTAT_COL)    ' статус Договора в SF
            ContrK = .Cells(i, 8)
            ContrId = ContractId(ContrK)
            If .Cells(i, 1) = 1 Then    ' В SF есть Орг.
                If ContrId <> "" Then   '==== Договор внесен в SF?
                    OppN = WorksheetFunction.VLookup(ContrK, _
                        Sheets(SFD).Range("B:Q"), 16, False)
                    If OppN = 0 Then
                        Call OppForDog(i)
                    Else
                '---- анализ существующего Проекта -------
                        OppIs = OppIs + 1
                        OppTyp = WorksheetFunction.VLookup(OppN, _
                            Sheets(SFopp).Range("B:R"), 17, False)
                        Vendor1C = .Cells(i, VENDOR1C_COL)
                        If Not IsSameVendor(OppTyp, Vendor1C, ContrK) And _
                                    Vendor1C <> "" Then
                            LogWr "     WARNING:  Поставщик1С = '" & Vendor1C & _
                                "' Тип... = " & OppTyp & _
                                " в OppN = " & OppN & " несовместимы! Договор " _
                                & ContrK
                        End If
                    End If
                Else                    '===== вносим новый Договор
                    NewContract .Cells(i, DOG1C_COL), .Cells(i, DOG1C_MAINDOG_COL), ContrK
                End If
            Else
'!!                NewAcc .Cells(i, DOG1CACC_COL)
            End If
        End With
    Next i

    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\Dogovor\"
    WriteCSV NewContract, "Dogovor.txt"
    WriteCSV NewContractLnk, "ContrUpd.txt"
    Shell "quotaDogovor.bat"
    Shell "quotaContUpd.bat"

    MsgBox "Найдено: " & vbCrLf & vbCrLf _
        & "   похожих организаций " & Fruitful _
        & vbCrLf & "   Проекты есть=" & OppIs _
        & vbCrLf & "  новых Проектов=" & OppNew & " отложено=" & NoOpp

    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\OppInsert\"
    WriteCSV O_NewOpp, "OppInsert.txt"
    Shell "quota2.bat"

    ModEnd DOG_SHEET
End Sub
Function OppForDog(iDog) As Boolean
'
' Есть ли подходящий Проект для связывания с ним Договора?
' Если есть - связываем (ContrLink); Если нет - создаем новый (NewOpp)
' Основная работа происходит с листом Договоров DOG_SHEET (=3), iDog - его строка
'   22.3.12

    Dim OppLnk As String        ' номер Проекта, с которым можно связать Договор
    Dim Acc, Dat, Sale, DogValue, DogCur As String  ' параметры нового Проекта
    Dim Stage, t As String      ' Этап и Тип нового Проекта по Договору
    Dim ContrK As String        ' Код Договора в форме Осн/Договор
    
    Dim t0, Tend
    
    t0 = Timer
        
    With Sheets(DOG_SHEET)
        Acc = .Cells(iDog, DOG1CACC_COL)
        Dat = .Cells(iDog, DOG1CDAT_COL)
        Sale = .Cells(iDog, DOG1CSALE_COL)
        DogCur = CurISO(.Cells(iDog, DOGCUR1C_COL))
        DogValue = .Cells(iDog, DOGSUM1C_COL) * CurRate(DogCur)
        ContrK = .Cells(iDog, DOGCOD_COL)
        t = OppT(.Cells(iDog, VENDOR1C_COL), Sale)
           
        OppLnk = IsOpp(Sale, Acc, t, DogValue, Dat, ContrK)
    
        If OppLnk = "" Then     ' есть подходящий Проект, куда войдет Договор?
                                    '---- нет - создаем новый Проект
            Stage = "70%-контракт на подписи у заказчика"
            If .Cells(iDog, DOGINVC_COL) <> 0 Then
                Stage = "80%-подписан контракт или счет в оплате"
            End If
            If .Cells(iDog, DOGPAID_COL) <> 0 Then
                Stage = "90%-первые деньги пришли на счет"
            End If
'!!            NewOpp Acc, ContrK, Dat, Sale, DogValue, DogCur, T, Stage
        Else                        '---- есть - связываем с ним Договор
            Call ContrLink(ContrK, OppLnk)
            LogWr "<L> Связь Договора '" & ContrK & "' с Проектом " & OppLnk
        End If
    End With
    Tend = (Timer - t0) * 1000
'    MsgBox "iDog = " & iDog & "Timer=" & Tend & " мсек"
'    If iDog > 100 Then Stop
End Function
Function OppT(V1C, Sale) As String
'
' возвращает Тип Проекта SF по имени Поставщика в Договоре 1С
'   3.2.12
'  20.3.12 - учет неоднозначности в таблице We "Тема&Вид деятельности"

    Dim S As Range
    OppT = V1C
    If OppT = "" Then Exit Function
    For Each S In Range("Тема_Вид_деятельности").Rows
        OppT = S.Cells(1, 1)
        If InStr(S.Cells(1, 2), V1C) <> 0 Then
            Select Case S.Cells(1, 3)
            Case "1:1", "1:M"
                Exit Function
            Case "M:1"
                If IsRightSale(Sale, OppT) Then Exit Function
                OppT = ""
            Case Else
                LogWr "FATAL ERROR OppT: нарушена таблица Wе Тема&Вид деятельности"
                StopSub
            End Select
        End If
    Next S
End Function
Sub DogOppLink()
'
' [*] проход по SFD и поиск подходящих Проектов для связи, Update Dogovor.csv
'   14.2.2012

    
    Dim i, j, Ndog As Integer
    Dim SaleDog, SaleOpp, OppN, OppType
    Dim DogDateStart, DogDateEnd, OppCloseDate As Date
    Dim Acc As String
    Dim ContrCode, Vendor1C As String

    Lines = ModStart(SFD, "DogOppLink: связь Договора с Проектом", True) - SFresLines

    CheckSheet SFD, Lines + 2, 3, SFcontrRepName
    CheckSheet SFopp, EOL_SFopp + 2, 1, SFoppRepName

    ClearSheet NewContractLnk, Range("HDR_ContrLnk")
    
'-- проход по листу SFD - по Договорам
    For i = 2 To Lines
        Progress i / Lines
        With Sheets(SFD)
            If Trim(.Cells(i, 1)) <> "" Then    ' работаем только с Договорами уже в SF
                Acc = .Cells(i, SFD_ACC1C_COL)
                SaleDog = .Cells(i, SFD_OWNER_COL)
                ContrCode = .Cells(i, SFD_COD_COL)
                Vendor1C = WorksheetFunction.VLookup(ContrCode, _
                    Sheets(DOG_SHEET).Range("H2:P77777"), 2, False)
                DogDateStart = .Cells(i, SFD_DATESTART_COL)
                DogDateEnd = .Cells(i, SFD_DATEEND_COL)
                If .Cells(i, SFD_OPPN_COL) = "" Then     '==== Договор связан с Проектом?
                '-- проход по SFopp - по Проектам           нет - связываем
                    For j = 2 To EOL_SFopp
                        With Sheets(SFopp)
                            SaleOpp = .Cells(j, SFOPP_SALE_COL)
                            OppN = .Cells(j, SFOPP_OPPN_COL)
                        ' для данного Договора выбираем только Проекты с той же Организацией
                        '       ..и с подходящим (или с тем же) Продавцом
                            If .Cells(j, SFOPP_ACC1C_COL) = Acc And _
                                    IsSameTeam(SaleDog, SaleOpp, OppN) Then
                                OppType = .Cells(j, SFOPP_TYP_COL)
                                OppCloseDate = .Cells(j, SFOPP_CLOSEDATE_COL)
                        ' проверка по  Поставщику и Датам Договора
                                If IsSameVendor(OppType, Vendor1C, ContrCode) And _
                                        OppCloseDate <= DogDateEnd And _
                                        OppCloseDate >= DogDateStart Then
                                    Call ContrLink(ContrCode, .Cells(j, SFOPP_OPPID_COL))
                                End If
                            End If
                        End With
                    Next j
                Else                                    '==== уже связан с Проектом
                    Dim DogPaidSum                          ' сумма платежей по Договору
                    DogPaidSum = WorksheetFunction.SumIf(Sheets(SF).Range("Q:Q"), _
                        .Cells(i, SDF_CONTRN_COL), Sheets(SF).Range("F:F"))
                    If Deviation(DogPaidSum, .Cells(i, SFD_VALUE_COL)) < 0.05 Then
                        Dim DogStatus
                        DogStatus = .Cells(i, SFD_STATUS_COL)
                        If DogPaidSum = 0 Then
                            LogWr " Договор не оплачен и его сумма =0"
                        ElseIf DogStatus <> DOG_STAT_CLOSED Then
                            LogWr "<!> Договор " & .Cells(i, SFD_COD_COL) _
                                & " ( SFD № = " & .Cells(i, SDF_CONTRN_COL) _
                                & ") надо ЗАКРЫТЬ"
                        End If
                    Else
                        i = i
                    End If
                End If
            End If
        End With
    Next i
    
'    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\OppInsert\"
'    WriteCSV O_NewOpp, "OppInsert.txt"
'    Shell "quota2.bat"
    ModEnd SFD
End Sub
Function Deviation(x, Y)
'
' возвращает относительную разницу Х и Y
'   15.3.12

    Const ErrVal = 999999
    If x <> 0 Then
        Deviation = Abs((x - Y) / x)
    ElseIf Y = 0 Then
        Deviation = 0
    Else
        Deviation = ErrVal
    End If
End Function
Function IsSameVendor(OppType, V1C, ContrCode) As Boolean
'
' возвращает True если Тип Проекта соответствует Поставщику в 1С со сверкой по We
'   5.3.12
' 17.3.12 - расширение LogWr

    Dim VendorSF As String  ' Тип&Вид деятельности по Проекту
    
    Dim OppTypeArr() As String
    Dim i
    
    IsSameVendor = False
    
' цикл по типам Проектов входящим в OppType -- в Типе их может быть несколько
    OppTypeArr = Split(OppType, ";")
    For i = 0 To UBound(OppTypeArr)
        VendorSF = ""
        On Error Resume Next
        VendorSF = WorksheetFunction.VLookup(Trim(OppTypeArr(i)), _
            Sheets(We).Range("Тема_Вид_деятельности"), 2, False)
        On Error GoTo 0
        If VendorSF = "" And V1C <> "" Then
            LogWr "     WARNING: [IsSameVendor] Неизвестный Тип&Вид = '" _
                & OppType & "' в SF, а в Договоре 1С '" & ContrCode _
                & "' Поставщик = '" & V1C & "'"
        End If
        If InStr(VendorSF, V1C) <> 0 Then GoTo Found
    Next i
    Exit Function
Found:
    IsSameVendor = True
End Function
