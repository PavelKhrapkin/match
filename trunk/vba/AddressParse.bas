Attribute VB_Name = "AddressParse"
'---------------------------------------------------------------------------------
' AddressParse  - анализ Адресов Организаций и слияние с внешними базами данных
'
' T TestAddr()        - Отладка AdrParse - разбор строки адреса
' - AdrParse(Address) - разбор строки почтового адреса
' - SeekDIC - подбираем подходящее значение по словарю в Range(DicRange) для Adr
' - adrCompRemove   - удалить компоненту адреса.
' - AdAccFr1C(Acc)  - записываем новую Организацию Acc из 1С в лист A_Acc для занесения в SF
' T testfillADic    - fillADic
' - fillADic()      -  заполнение hash-таблицы по листу A_Dic (слова употребляемые в именах предприятий)
' T testfillSFacc   - тест fillSFacc
' - fillSFacc       - заполнение hash-таблицы по листу SFacc
' T testfillAcc1C
' - fillAcc1C       - заполнить hash по Acc1C - словарь предприятий 1С
'   20.5.2012   A.Пасс
    
Option Explicit

' hash таблица для AdAccFr1C - имена клиентов
Dim accHTbl(0 To 5000) As String
Dim accVTbl(0 To 5000) As String

Dim DicReadFlag As Boolean              ' статически инициализируется как False
Dim aDicKey(0 To 9999) As String        ' hash словарь
Dim aDicVal(0 To 9999) As String

Dim SFaccKey(0 To 4999) As String       ' hash словарь предприятий в 1С
Dim SFaccVal(0 To 4999) As String
Dim Acc1CKey(0 To 4999) As String       ' hash словарь предприятий в SF
Dim Acc1CVal(0 To 4999) As String
 
Sub TestAddr()
'
' T testAddr() - Отладка AdrParse - разбор строки адреса
'   14.5.12 А.Пасс

    Dim A1(0 To 20) As PostAddr

    A1(18) = AdrParse("196105. г. Санкт-ПетербургX.Витебский пр., д.11     ")
    A1(16) = AdrParse(" г. Санкт-ПЕтербург. ул. Комсомола д. 1-3, тел. / факс  592-43-60")
    A1(17) = AdrParse("198323, ЛЕНИНГРАДСКАЯ ОБЛАСТЬ, ЛОМОНОСОВСКИЙ РАЙОН, Д МАЛОЕ КАРЛИНО, ШОССЕ ВОЛХОНСКОЕ, 7")
    A1(15) = AdrParse(" Московская обл. Ломоносовский р Д.  Малое Карлино ул. Калинина, д.1 143026")
    
    ' проверяем с запятыми
    
    A1(1) = AdrParse("198323, ЛЕНИНГРАДСКАЯ ОБЛАСТЬ, ЛОМОНОСОВСКИЙ РАЙОН, Д МАЛОЕ КАРЛИНО, ШОССЕ ВОЛХОНСКОЕ, 7")
    A1(2) = AdrParse("195009, г. Санкт-ПЕтербург, ул. Комсомола д. 1-3, тел. / факс  592-43-60")
    A1(3) = AdrParse("128234, Россия, Плеханова 25 кв 8, Москва")
    A1(4) = AdrParse("128234, Россия, Плеханова 25 кв 8, Н-ск")
    A1(5) = AdrParse("128234, Россия, Плеханова 25 кв 8, Вологда")
    A1(6) = AdrParse("AAA")                                                                                         ' ERR
    A1(7) = AdrParse("")                                                                                            ' ERR
    A1(8) = AdrParse("143026, Московская обл., Одинцовский р-он,р.п.Новоивановское, ул. Калинина, д.1         ")
    A1(9) = AdrParse("143026 Московская обл., Одинцовский р-он,р.п.Новоивановское, ул. Калинина, д.1         ")
    A1(10) = AdrParse(" 143026, Московская обл., Одинцовский р-он,р.п.Новоивановское, ул. Калинина, д.1         ")
    A1(11) = AdrParse(" Московская обл., Одинцовский р-он,р.п.Новоивановское, ул. Калинина, д.1 , 143026")
    A1(12) = AdrParse("    +143026")                                                                                ' ERR

    ' проверяем без запятых
    
    A1(13) = AdrParse(" Московская обл. Одинцовский р-он р.п.Новоивановское ул. Калинина, д.1 143026")
    A1(14) = AdrParse(" Московская обл. Одинцовский р-он р.п. Новоивановское ул. Калинина, д.1 143026")
    A1(15) = AdrParse(" Московская обл. Ломоносовский р Д.  Малое Карлино ул. Калинина, д.1 143026")
    
    Dim clIndx As Integer
    Dim addrToParse As String
    
    ' проверяем по базе данных 1С
'
    EOL_AdAcc = 1
    For clIndx = 2 To 9999
        If Sheets(Acc1C).Cells(clIndx, A1C_NAME_COL) = "" Then GoTo endCLoop   ' пустое имя - конец цикла
        addrToParse = Sheets(Acc1C).Cells(clIndx, A1C_ADR_COL)
        If addrToParse <> "" Then
                
            clIndx = clIndx                     ' для отладки
            If A1(0).ErrFlag Then
                clIndx = clIndx                     ' для отладки
            Else
            
                EOL_AdAcc = EOL_AdAcc + 1
'If EOL_AdAcc = 9 Then
'            ' повторяем парсинг для отладки
'            A1(20) = AdrParse(addrToParse)
'End If
                With Sheets(A_Acc)
                    .Cells(EOL_AdAcc, ADACC_NAME_COL) = Sheets(Acc1C).Cells(clIndx, A1C_NAME_COL)
                    .Cells(EOL_AdAcc, ADACC_1CNAME_COL) = Sheets(Acc1C).Cells(clIndx, A1C_NAME_COL)
                    .Cells(EOL_AdAcc, ADACC_INN_COL) = Sheets(Acc1C).Cells(clIndx, A1C_INN_COL)
                    .Cells(EOL_AdAcc, ADACC_INDEX_COL) = A1(0).PostIndex
                    .Cells(EOL_AdAcc, ADACC_CITY_COL) = A1(0).City
                    .Cells(EOL_AdAcc, ADACC_STREET_COL) = A1(0).Street
                    .Cells(EOL_AdAcc, ADACC_COUNTRY_COL) = A1(0).Country
                    .Cells(EOL_AdAcc, ADACC_STATE_COL) = A1(0).State
                    .Cells(EOL_AdAcc, ADACC_TEL_COL) = Sheets(Acc1C).Cells(clIndx, A1C_TEL_COL)
                End With
                
            End If
        End If
    Next clIndx

endCLoop:

End Sub
Function AdrParse(ByVal Address As String) As PostAddr
'
' AdrParse(Address)   - разбираем строку почтового адреса Address
'   11.5.12

'    AdrParse = AdrParse(Address, True)
''    If AdrParse.ErrFlag Then
''        AdrParse = AdrParse(Address, False)    ' если не распарсили - пробуем без запятых
''    End If
'End Function
'
'
'Function AdrParse(ByVal Address As String, commaSplit As Boolean) As PostAddr

' - AdrParse(Address) - разбор строки почтового адреса с указанным типом split
'       (по запятым, если true, по пробелам, если false)
'   11.5.12

    Dim Adr() As String
    Dim i As Integer, j As Integer
    Dim lastIndxPos As Integer
    
'    Const addrExtraWrds = "г республика респ дер рп п "
    
    AdrParse.ErrFlag = False

' --- ищем почтовый индекс

    Address = Address + ","             ' чтобы упростить проверку на цифру в конце строки
    AdrParse.PostIndex = ""
    lastIndxPos = -100                  ' чтобы проверка длины заведомо не сработала
    For i = 1 To Len(Address)
        If Not IsNumeric(Mid(Address, i, 1)) Then
            If i - lastIndxPos = 6 Then ' конец последовательности цифр. проверяем длину
                AdrParse.PostIndex = Mid(Address, lastIndxPos, 6)
                Mid(Address, lastIndxPos, 6) = "      "
                GoTo split              ' индекс найден
            End If
            lastIndxPos = -100
        Else
            If lastIndxPos < 0 Then lastIndxPos = i     ' цифра. Она первая в серии?
        End If
    Next i
    
split:
    Adr = split(Compressor(Address), ",")
    
'--- ищем город, область, страну
    AdrParse.City = SeekDIC("DIC_CityNames", Adr)
    AdrParse.Country = SeekDIC("DIC_CtryNames", Adr)
    AdrParse.State = SeekDIC("DIC_Area", Adr)
    
    If AdrParse.Country = "" Then AdrParse.Country = "Россия"

'--- формируем поле улица
    Dim Street As String
    Dim x As String, curr As String
    
    ' включаем неудаленные компоненты в Street через запятую
    Street = ""
    For i = LBound(Adr) To UBound(Adr)
' пропускаем пустые строки и одинокие точки
        x = Trim$(Adr(i))
        For j = 1 To Len(x)
            curr = Mid(x, j, 1)
            ' ищем значимые символы, найдя - выходим
            If curr <> " " And curr <> "." Then GoTo breakL     ' оператор break
            Mid(x, j, 1) = " "                                   ' удаляем лидирующие точки и заодно пробелы
        Next j
breakL:
        If x <> "" Then
            If Street <> "" Then Street = Street & ", "
            Street = Street & x
        End If
    Next i
    
    AdrParse.Street = Compressor(Street)
    
'--- формируем ErrFlag (state можно опускать)
    If AdrParse.City = "" Or InStr(AdrParse.City, "?") <> 0 _
            Or AdrParse.PostIndex = "" Or AdrParse.Street = "" Then
        AdrParse.ErrFlag = True
    End If
    
End Function
Function SeekDIC(DicRange, Adr) As String
'
' SeekDIC - подбираем подходящее значение по словарю в Range(DicRange) для Adr
'   14.5.12


    Dim sWe As Range
    Dim wrds() As String
    Dim wrdsArea() As String
    Dim SeekWord As String, pattern As String, PatternArea As String, curr As String, SeekArea As String
    Dim i As Integer, j As Integer, k As Integer, kArea As Integer
    Dim wrdPos As Integer
    
    SeekDIC = ""
    For i = LBound(Adr) To UBound(Adr)      ' цикл по компонентам адреса
        If Adr(i) <> "" Then
            SeekWord = LCase$(Adr(i))
            For Each sWe In Range(DicRange).Rows    ' цикл по Range (один из: DIC_CityNames, DIC_CtryNames, DIC_Area)
                wrds = split(sWe.Cells(1, 1), ",")
                For k = LBound(wrds) To UBound(wrds)    ' цикл по словам кол. 1 (допустимые написания, начала)
                    pattern = Trim(LCase$(wrds(k)))     ' Trim позволяет писать там пробелы вокруг запятых
                    wrdPos = InStr(SeekWord, pattern)
                    If wrdPos <> 0 Then
                        
                        ' найденную компоненту удаляем из массива
                        If DicRange = "DIC_CityNames" Then
                            adrCompRemove Adr(i), wrdPos, pattern
                        Else
                            Adr(i) = ""
                        End If
                        SeekDIC = wrds(LBound(wrds))    ' найден целевой вариант
                        
                        ' ищем район, если попросили в кол. 2 табл. "We"
                        ' район вставляется в результат перед городом, через запятую

                        If sWe.Cells(1, 2) <> "" Then
                            For j = LBound(Adr) To UBound(Adr)     ' цикл по компонентам адреса
                                If Adr(j) <> "" Then
                                    SeekArea = LCase$(Adr(j))
                                    wrdsArea = split(sWe.Cells(1, 2), ",")
                                    For kArea = LBound(wrdsArea) To UBound(wrdsArea) ' цикл по словам кол. 2 (допустимые написания, начала)
                                        PatternArea = Trim(LCase$(wrdsArea(kArea)))
                                        wrdPos = InStr(SeekArea, LCase$(wrdsArea(kArea)))
                                        If wrdPos <> 0 Then
                                            adrCompRemove Adr(j), wrdPos, PatternArea
                                            SeekDIC = wrdsArea(LBound(wrdsArea)) & ", " & SeekDIC   ' найден целевой вариант
                                            Exit Function
                                        End If
                                    Next kArea
                                End If
                            Next j
                            
                            SeekDIC = wrdsArea(LBound(wrdsArea)) & ", ?"        ' район должен найтись, если запрошен в "We"
                            
                        End If
                            
                        Exit Function
                    End If
                Next k
            Next sWe
        End If
    Next i
    

End Function
Sub adrCompRemove(AdrComp, wrdPos, pattern)

'   adrCompRemove - удалить компоненту адреса.
'       делается попытка забить слово пробелами
'   14.05.2012

    Dim patternEnd As Integer, i As Integer, lastpos As Integer
    Dim curr As String, wspflag As Boolean, wrd As String
    
'    If commaSplit Then
'        AdrComp = ""
'    Else
        ' забиваем пробелами целые слова общей длиной не меньше чем Len(Pattern)
        patternEnd = wrdPos + Len(pattern) - 1  ' последняя позиция слова
        lastpos = 0                             ' на случай, если не присвоим
        For i = wrdPos To Len(AdrComp)
            curr = Mid(AdrComp, i, 1)
        
            If i > patternEnd And (curr = " " Or curr = ".") Then GoTo endLoop
            lastpos = i
'
'            If i <= patternEnd Or (curr <> " " And curr <> "," And curr <> ".") Then
'                Mid(AdrComp, i, 1) = " "
'            Else
'                GoTo endLoop
'            End If
        Next i
endLoop:
        If wrdPos <= lastpos Then Mid(AdrComp, wrdPos, lastpos - wrdPos + 1) = "                                        "
       
' найти предшествующее слово и, если оно запрещенное (напр., "д"), удалить его
        If wrdPos > 1 Then
            wspflag = True
            i = wrdPos - 1
            Do
                curr = Mid(AdrComp, i, 1)
                If wspflag Then
                    ' пропускаем пробелы и точки справа
                    If curr <> " " Or curr <> "." Then wspflag = False
                Else
                    ' ищем терминирующий пробел
                    If curr = " " Then
                        i = i + 1       ' нашли, отходим вправо на 1 позицию и выходим
                        GoTo remPrefix
                    End If
                End If
                If i = 1 Then
                    GoTo remPrefix      ' начало строки, выходим
                Else
                    i = i - 1           ' двигаемся дальше влево
                End If
            Loop
remPrefix:
' Добавляется отдельное слово в Street.
' снять хвостовые/нач пробелы, понизить регистр, убрать точки и проверить - запрещено?
            wrd = Replace(LCase$(Trim(Mid(AdrComp, i, wrdPos - i))), ".", "") & " "
' слова, запрещенные к включению в Street, заданы литералом. Пробел должен быть после каждого слова,
' все должно быть в нижнем регистре и без точек.
            If InStr("г гор республика респ д дер рп п пос", wrd) <> 0 Then
                Mid(AdrComp, i, wrdPos - i) = "                                        "
            End If
        End If
        
'    End If
End Sub
Sub testAddAcc1C()

    ClearSheet A_Acc, Range("HDR_AdAcc")
    EOL_Acc1C = EOL(Acc1C) - ACC1C_RES
    
    hashInit accHTbl, accVTbl
    
    AdAccFr1C ("ЗАО ""Актуальная геология""")
    AdAccFr1C ("ЗАО ""Актуальная геология""")       ' проверка на дубликат
    AdAccFr1C ("Марина Ефимовна     ")

End Sub
Sub testAdAccFr1C()

    AdAccFr1C "ООО ""ГеоИнжиниринг"""

    AdAccFr1C "ЦМТиХО тыла ГУ МВД России по г.Санкт-Петербургу и Ленинградской области"
    AdAccFr1C "xxxxxxxxxxx"     ' нет в 1С
    AdAccFr1C "Чегевара"        ' плохой адрес

End Sub
Sub iniAdAccFr1C()
'   теперь инциализируется в AdAccFr1C
'    hashInit accHTbl, accVTbl
End Sub

Sub AdAccFr1C(acc)
'
' записываем новую Организацию Acc из 1С в лист A_Acc для занесения в SF
'   16.4.12

    Dim INN As String, Index As String, Street As String
    Dim Country As String, State As String, tel As String
    Dim Adr() As String
    
    Dim Addr As PostAddr, addrString As String
    
    Dim accWords() As String
    Dim accIndxStr As String, accIndx As Long
    
    Dim i, j As Integer
    Dim Found As Boolean
    Found = False
    
    If Not DicReadFlag Then        ' предполагаем, что DicReadFlag статически инициализирован как False
        hashInit accHTbl, accVTbl
        fillADic
        fillSFacc
        fillAcc1C
        DicReadFlag = True
    End If
    
' если новая организация уже есть в листе A_Acc - ничего не делаем
    If hashGet(accHTbl, accVTbl, acc) <> "$" Then GoTo ExitSub
    
' Ищем каждое слово из account в A_Dic. Если находим, не заносим, но пишем в Log
    accWords = split(acc, " ")
    For i = LBound(accWords) To UBound(accWords)
        If hashGet(aDicKey, aDicVal, Trim$(accWords(i))) <> "$" Then
            LogWr "<!> Account '" & acc & "'" _
                & "' содержит слово '" & accWords(i) & "', имеющееся в словаре."
            GoTo ExitSub
        End If
    Next i
        
' тест: есть ли она в SF
    If hashGet(SFaccKey, SFaccVal, acc) <> "$" Then
        LogWr "<!> Account '" & acc & "'" & "' присутствует в поле 'Имя Организации в 1С' "
    End If
                
' дедупликация - повторно организацию не заводим
    For i = 2 To EOL_AdAcc
        If acc = Sheets(A_Acc).Cells(i, ADACC_NAME_COL) Then GoTo ExitSub
    Next i
    
' ищем account в словаре 1С, чтобы извлечь адрес и др. данные (см. fillAcc1C)

    accIndxStr = hashGet(Acc1CKey, Acc1CVal, acc)
    If accIndxStr = "$" Then GoTo ExitSub               ' не нашли в словаре 1С
    accIndx = Val(accIndxStr)                           ' получить целое
    
    With Sheets(Acc1C)
       
        INN = .Cells(accIndx, A1C_INN_COL)
        If INN <> "" Then INN = split(INN, "/")(0)
            
''''''''''''''''''''''''''''''''''
        addrString = .Cells(accIndx, A1C_ADR_COL)
        If addrString <> "" Then
            Addr = AdrParse(addrString)
            If Addr.ErrFlag Then
                LogWr "<!> Ошибка разбора адреса для '" & acc & "'" _
                    & "; адрес '" & Trim(addrString) & "'"
            Else
                GoTo FoundAdr
            End If
  ''''''''''''''''''''''''''''''''''
        End If
    End With
    GoTo ExitSub            ' нет в 1С - выходим
    
    
            '        For i = 2 To EOL_Acc1C
            '    For i = 1 To EOL_SFacc
            '        ' проверяем по 1С имени
            '        If acc = Sheets(SFacc).Cells(i, SFACC_ACC1C_COL) Then
            '            LogWr "<!> Account '" & acc & "'" & "' присутствует в поле 'Имя Организации в 1С' "
            '        End If
            '    Next i
    

FoundAdr:
    EOL_AdAcc = EOL_AdAcc + 1
    
    With Sheets(A_Acc)
        .Cells(EOL_AdAcc, ADACC_NAME_COL) = acc
        .Cells(EOL_AdAcc, ADACC_1CNAME_COL) = acc
        .Cells(EOL_AdAcc, ADACC_INN_COL) = INN
        .Cells(EOL_AdAcc, ADACC_INDEX_COL) = Addr.PostIndex
        .Cells(EOL_AdAcc, ADACC_CITY_COL) = Addr.City
        .Cells(EOL_AdAcc, ADACC_STREET_COL) = Addr.Street
        .Cells(EOL_AdAcc, ADACC_COUNTRY_COL) = Addr.Country
        .Cells(EOL_AdAcc, ADACC_STATE_COL) = Addr.State
        .Cells(EOL_AdAcc, ADACC_TEL_COL) = Sheets(Acc1C).Cells(accIndx, A1C_TEL_COL)    ' phone#
'        .Cells(EOL_AdAcc, ADACC_CONT_COL) = Sheets(Acc1C).Cells(accIndx, A1C_CONT_COL)  ' Контакт
    End With

' Результат сохраняем в hash таблице (accHTbl,accVTbl)
    hashSet accHTbl, accVTbl, acc, ""
    
ExitSub:
End Sub
Sub testfillADic()
    fillADic
End Sub
Sub fillADic()
' заполнение hash-таблицы по листу A_Dic (слова употребляемые в именах предприятий)
' 18.05.12

' используется только ключ, value не заполняется

    Dim i As Integer, x As String
    
    hashInit aDicKey, aDicVal
    For i = 2 To 9999
'If i = 4150 Then
'i = i
'End If
        x = Sheets(A_Dic).Cells(i, 1)
        If x = "" Then
            GoTo ExitSub
        End If
        hashSet aDicKey, aDicVal, x, ""
    Next i
ExitSub:
End Sub
Sub testFillSFacc()
' Т тест fillSFacc
'   19.5.2012
    fillSFacc
End Sub
Sub fillSFacc()
' заполнение hash-таблицы по листу SFacc
' 18.05.12

' используется только ключ, value не заполняется

    Dim i As Long, x As String
    
    hashInit SFaccKey, SFaccVal
    
    Dim ef As Long
    ef = EOL_Acc1C
    If ef = 0 Then ef = 9999                    ' ТОЛЬКО ДЛЯ ОТЛАДКИ

    For i = 2 To ef
        If Sheets(SFacc).Cells(i, 1) <> "" Then                    ' игнорируем пустые
            x = Sheets(SFacc).Cells(i, SFACC_ACC1C_COL)
            If i = 3589 Then
                i = i
            End If
            If ef = 9999 Then                       ' ТОЛЬКО ДЛЯ ОТЛАДКИ
                If Sheets(SFacc).Cells(i, 1) = "SFacc" Then GoTo ExitSub    ' ТОЛЬКО ДЛЯ ОТЛАДКИ
            End If                                  ' ТОЛЬКО ДЛЯ ОТЛАДКИ
        End If
    Next i
ExitSub:
End Sub
Sub testfillAcc1C()
' Т тест fillAcc1C
'   19.5.2012
    fillSFacc
End Sub
Sub fillAcc1C()

' заполнить hash по Acc1C - словарь предприятий 1С
'   19.5.2012

    Dim i As Integer, x As String
    
    Dim ef As Long
    ef = EOL_SFacc
    If ef = 0 Then ef = 9999        ' ТОЛЬКО ДЛЯ ОТЛАДКИ

    hashInit Acc1CKey, Acc1CVal
    For i = 1 To ef
        x = Sheets(Acc1C).Cells(i, A1C_NAME_COL)
        If x = "" Then GoTo ExitSub
        If ef = 9999 Then                       ' ТОЛЬКО ДЛЯ ОТЛАДКИ
            If x = "SFacc" Then GoTo ExitSub    ' ТОЛЬКО ДЛЯ ОТЛАДКИ
        End If                                  ' ТОЛЬКО ДЛЯ ОТЛАДКИ
        ' записываем в hash исключая дубликаты
        
If x = "Чегевара" Then
i = i
End If
        
        If hashGet(Acc1CKey, Acc1CVal, x) = "$" Then hashSet Acc1CKey, Acc1CVal, x, i
    Next i
ExitSub:
End Sub
