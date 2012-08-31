Attribute VB_Name = "ContrAnalitics"
'-----------------------------------------------------------------------------
' ContrAnalitics    - анализ Контрактов и относящихся к ним Проектов
' [*] ContrPass     - проход по листу Договоров (3) и создание новых Проектов
'  -  OppT(V1C)     - возвращает тип Проекта по полю Поставщик в Договоре 1С
' [*] DogOppLink    - проход по SFD и поиск подходящих Проектов для связи
'  -  IsSameVendor(OppType, V1C, ContrCode)    - возвращает True, если Тема&Вид
'                           Проекта OppType соответствует Поставщику по Договору в 1С
'   19.8.2012

Option Explicit
Sub NewContr()
'
' - NewContr()  - просмотр Договоров 1С для занесения в SF новых через DL
' 18.8.2012

    Dim Dog As TOCmatch
    Dim i As Long
    
    Dog = GetRep(DOG_SHEET)
    NewSheet NewContract

    With DB_1C.Sheets(DOG_SHEET)
        For i = 2 To Dog.EOL
            Progress i / Dog.EOL
'            If ExRespond = False Then Exit For
            If .Cells(i, DOGIDSF_COL) = "" And .Cells(i, DOGISACC_COL) <> "" Then
                WrNewSheet NewContract, DB_1C.Sheets(DOG_SHEET), i
            End If
'         - WrNewSheet(SheetNew, SheetDB, DB_Line, ToWriteCols, HDR_FormName)
'            NewContract .Cells(i, DOG1C_COL), .Cells(i, DOG1C_MAINDOG_COL), ContrK
        Next i
    End With
End Sub
Sub NewSheet(SheetName)
'
' - NewSheet(SheetName, HDRform) - создает новый лист SheetName
'       Название шапки нового листа берется из названия SheetName,
'       а ширина колонок шапки- из третьей cтроки формы
' 19.8.12

    Dim HDRform As String
    Dim i As Long
    
    HDRform = "HDR_" & SheetName
    
    With DB_MATCH
        On Error Resume Next
        .Sheets(SheetName).Delete
        .Sheets.Add After:=.Sheets(.Sheets.count)
        .Sheets(.Sheets.count).Name = SheetName
        On Error GoTo 0
        With .Sheets(SheetName)
            .Tab.Color = rgbLightBlue
            .Activate
            For i = 1 To Range(HDRform).Columns.count
                Range(HDRform).Columns(i).Copy Destination:=.Cells(1, i)
                .Columns(i).ColumnWidth = .Cells(3, i)
            Next i
            .Rows(6).Delete
            .Rows(5).Delete
            .Rows(4).Delete
            .Rows(3).Delete
            .Rows(2).Delete
        End With
    End With
End Sub
Sub WrNewSheet(SheetNew, SheetDB, DB_Line)
'
' - WrNewSheet(SheetNew, SheetDB, DB_Line) - записывает новый рекорд в лист SheetNew
'                                            из строки DB_Line листа SheetDB
'   * Имя и Параметры для обработки передаются в Адаптер в виде текстовых строк.
'     Эти строки хранятся в Range с именем "HDR_" & SheetNew в Forms или Headers
'   * Обращение к Адаптеру имеет вид <ИмяАдаптера>/<Пар1>,<Пар2>...
'   * В строке формы под Адаптером можно указать параметры во внешних Документах
' 29.8.2012

    Dim P As Range
    Dim iNewLine As Long    '= номер строки в SheetNew
    Dim i As Long
    Dim X As String         '= обрабатываемое значение в SheetDB
    Dim Y As String         '= результат работы Адаптера
    Dim IsErr As Boolean    '=True если Адаптер обнаружил ошибку
    
    iNewLine = EOL(SheetNew, DB_MATCH) + 1

    With DB_MATCH.Sheets(SheetNew)
        .Activate
        Set P = Range("HDR_" & SheetNew)
        For i = 1 To P.Columns.count
            X = SheetDB.Cells(DB_Line, P.Cells(4, i))
            
            Y = Adapter(P.Cells(5, i), X, P.Cells(6, i), IsErr)
            
            If IsErr Then
                .Rows(iNewLine).Delete
                Exit For
            Else
                .Cells(iNewLine, i) = Y
            End If
        Next i
    End With
End Sub
Function Adapter(Request, ByVal X, F_rqst, IsErr) As String
'
' - Adater(Request, X, F_rqst) - обрабатывает X с помощью Адаптера Request
'                                с внешними данными в Документе F_rqst
' 29.8.12

    Dim FF() As String, Tmp() As String, Cols() As String
    Dim Doc As String, C1 As Long, C2 As Long, Rng As Range
    Dim F() As String
    Dim i As Long, Par() As String
    
    IsErr = False
    
'--- разбор строки Адаптера вида <Имя>/C1,C2,C3...
    Dim AdapterName As String
    AdapterName = ""
    If Request <> "" Then
        Tmp = Split(Request, "/")
        AdapterName = Tmp(0)
        If InStr(Request, "/") <> 0 Then Par = Split(Tmp(1), ",")
    End If

'========== препроцессинг Адаптера =========
    Select Case AdapterName
    Case "MainContract":
        X = Trim(Replace(X, "Договор", ""))
    End Select
    
'--- FETCH разбор строки параметров из Документов вида <Doc1>/C1:C2,<Doc2>/C1:C2,...
    If F_rqst <> "" Then
        
        FF = Split(F_rqst, ",")
        For i = LBound(FF) To UBound(FF)
            Tmp = Split(FF(i), "/")
            Doc = Tmp(0)
            Cols = Split(Tmp(1), ":")
            C1 = Cols(0): C2 = Cols(1)
            GetRep Doc
            Set Rng = Workbooks(RepTOC.RepFile).Sheets(RepTOC.SheetN) _
                .Range(Columns(C1), Columns(C2))
            Dim S As String
            S = ""
            On Error Resume Next
            S = WorksheetFunction.VLookup(X, Rng, C2 - C1 + 1, False)
            On Error GoTo 0
            If S = "" Then
                ErrMsg WARNING, "Адаптер> ссылка " & F_rqst _
                    & "(" & X & ") не работает, результат <пусто>"
                IsErr = True
                Exit Function
            Else
                X = S
            End If
        Next i
    End If


'******* выполнение Адаптера с параметрами Par ******
    Select Case AdapterName
    Case "", "MainContract": Adapter = X
    Case "Мы", "Продавец_в_SF":
        Adapter = WorksheetFunction.VLookup(X, Range(AdapterName), Par(0), False)
    Case "Dec": Adapter = Dec(X)
    Case "CurISO": Adapter = CurISO(X)
    Case "CurRate": Adapter = CurRate(CurISO(X))
    Case "Дата": Adapter = DDMMYYYY(X)
    Case Else
        ErrMsg FATAL_ERR, "Adapter> Не существует " & AdapterName
    End Select
End Function
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
Ex:
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
Function Deviation(X, Y)
'
' возвращает относительную разницу Х и Y
'   15.3.12

    Const ErrVal = 999999
    If X <> 0 Then
        Deviation = Abs((X - Y) / X)
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
