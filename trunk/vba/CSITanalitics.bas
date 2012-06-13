 Attribute VB_Name = "CSITanalitics"
'----------------------------------------------------------
' Модуль анализа информации из CSIT
'   Пасс А.     13.6.12
' - CSIT_MS_Clear()                             - очистка состояния лидов MS_CSIT
' - CSIT_MS_lead()                              - анализ имен организаций - лидов Microsoft CSIT
' - SFPostAddr(indx As Long, SFacc As String)   - Стандартное представление почтового адреса

Option Explicit

Sub CSIT_MS_Clear()
' очистка состояния лидов MS_CSIT
'   31.05.12

    Dim i As Long
    EOL_CSIT_MS = EOL(CSIT_MS)
    For i = 7 To EOL_CSIT_MS                ' цикл по MS предприятиям
        Sheets(CSIT_MS).Cells(i, CSIT_MS_IDSF_COL) = ""
    Next i

End Sub

Sub CSIT_MS_lead()
'
' анализ имен организаций - лидов Microsoft CSIT
'   31.05.12

' процедура ставит * в строке предприятия, если не найдено соответствий SFacc,
'                  "X", если найдено, но оператор не подтверждает ни одного,
'                  список ID, разделенных пробелами, если оператор их подтверждает.
' анализируются только строки, содержащие "". Для приведения в исходное состояние используется CSIT_MS_Clear()

    Const Doing = "Анализ названий организаций - лидов MS_CSIT"
    ModStart CSIT_MS, Doing
    ProgressForm.Show vbModeless
    ProgressForm.ProgressLabel.Caption = Doing
    LogWr ""
    LogWr Doing
    ExRespond = True
    
    EOL_CSIT_MS = EOL(CSIT_MS)
    EOL_SFacc = EOL(SFacc) - SFresLines

    CheckSheet CSIT_MS, 4, 2, CSIT_MS_STAMP
    CheckSheet Acc1C, 1, 5, "Название фирмы"
    CheckSheet SFacc, EOL_SFacc + 2, 1, SFaccRepName
'---------- проход по CSIT_MS ------------------------
    Dim i As Long, j As Long, k As Long
    Fruitful = 0

    Dim SFname As String, SeekWords() As String, SNumsS() As String, SNums As Long, AccId As String
    
    ' Hash таблица содержащая словa из названий предприятий и номер строки в SFacc (ключ - компонента имени, слово)
    Dim SFAccKTbl(0 To 9999) As String, SFAccVTbl(0 To 9999) As String
    hashInit SFAccKTbl, SFAccVTbl    ' делаем таблицу пустой
    
    Dim curAcc As String, SFAccNums As String, SFId As String
    Dim Msg As String, MSG2 As String, Client As String, Respond As Long
    
    With Sheets(SFacc)
        For i = 2 To EOL_SFacc
            ' имя организации -> id
'            If i = 4 Then
'                i = i
'            End If
            SFname = RemIgnored(LCase$(.Cells(i, SFACC_ACCNAME_COL)))
            If Trim(SFname) = "" Then
                Msg = "Только игнорируемые слова в имени: '" & .Cells(i, SFACC_ACCNAME_COL) & "'. "
                LogWr Msg
'                Respond = MsgBox(MSG & vbCrLf & vbCrLf & " Тем не менее включаем?", vbYesNo)
'                If Respond <> vbYes Then GoTo NextI
                GoTo NextI      'ВРЕМЕННО!!!
                SeekWords = split(LCase$(.Cells(i, SFACC_ACCNAME_COL)), "$")        ' включаем как одно слово
            Else
                SeekWords = split(SFname, " ")                                      ' включаем слова по отдельности, без игнорируемых
            End If
            ' рассматриваем имя организации как слова и для каждого слова добавляем в таблицу номер строки SFacc
            For j = 0 To UBound(SeekWords)
                curAcc = hashGet(SFAccKTbl, SFAccVTbl, SeekWords(j))
                If curAcc <> "$" Then
                    curAcc = curAcc + " "
                Else
                    curAcc = ""
                End If
                hashSet SFAccKTbl, SFAccVTbl, SeekWords(j), curAcc + Trim(Str(i))     ' номер строки в таблице
            Next j
NextI:
        Next i
    End With
    
    Dim compNum As Long   ' компоненты имени предприятия в SF (array, index)
    Dim MSName As String                    ' имя предприятия в MS
    Dim SFnComps() As String, sfn As Long   ' компоненты имени предприятия в SF (array, index)
    Dim CompSNums(1 To 100) As Long         ' номерa компонент имени
    
    ' ищем компоненты из MS имени в hash - таблице компонент SF - имен
    
    For i = 7 To EOL_CSIT_MS                ' цикл по MS предприятиям
        Progress (i - 7) / EOL_CSIT_MS
        If ExRespond = False Then GoTo BreakForI
        AccId = "*"                         ' если ничего не найдем - поставим в таблицу
        With Sheets(CSIT_MS)
            MSName = .Cells(i, CSIT_MS_NAME_COL)
            If MSName <> "" And Trim(.Cells(i, CSIT_MS_IDSF_COL)) = "" Then         ' если не пусто - уже обрабатывали
            
                ' обработка одного MS предприятия
                
                Client = .Cells(i, CSIT_MS_NAME_COL)                                ' Формируем диалог - включаем MS account
                Msg = "CSIT_MS имя:" + "'" & Client & "';" _
                    + vbCrLf + "Адрес: " + .Cells(i, CSIT_MS_ADDR_COL) _
                    + vbCrLf + vbCrLf + "---- Возможные SF имена ----"
                MSG2 = ""
                SeekWords = split(RemIgnored(LCase$(MSName)), " ")                  ' разбиваем MS имя на компоненты
                compNum = 0
                For j = 0 To UBound(SeekWords)
                
                        ' строка SF номеров из hash извлеченная по компоненте MS-имени (разделитель - пробел)
                    SFAccNums = hashGet(SFAccKTbl, SFAccVTbl, SeekWords(j))
                    If SFAccNums <> "$" Then                                        ' в hash что-то есть (номера в таблице SFacc, таблице SF - accounts)?
                
                        MSG2 = MSG2 + vbCrLf + vbCrLf + "       КЛЮЧЕВОЕ СЛОВО: " + SeekWords(j)
                        SNumsS = split(SFAccNums, " ")
                        For k = 0 To UBound(SNumsS)                                 ' цикл по номерам SFacc
                            SNums = CInt(SNumsS(k))
                            SFnComps = split(LCase(RemIgnored(Sheets(SFacc).Cells(SNums, SFACC_ACCNAME_COL))), " ")
                            For sfn = 0 To UBound(SFnComps)                         ' цикл по компонентам SF-имени
                            
                                If SFnComps(sfn) = SeekWords(j) Then
                                    SFname = Sheets(SFacc).Cells(SNums, SFACC_ACCNAME_COL)
                                    compNum = compNum + 1
                                    MSG2 = MSG2 + vbCrLf + vbCrLf + "            " + Format(compNum) + ".     '" + "'" & SFname & "'" _
                                        & vbCrLf & "                 Адрес: " & SFPostAddr(SNums, SFacc)
                                        
                                    CompSNums(compNum) = SNums      ' запомнить номер строки в SFacc
                                    
                                End If
                                                            
                            Next sfn
                        Next k
                    End If
                Next j

' Текст подготовлен. Запускаем диалог.

                If MSG2 <> "" Then
                    Do
                        SFaccountForm.TextBox2.value = Msg + MSG2       ' основной текст
                        SFaccountForm.TextBox1.value = ""               ' исходное значение номера - пусто
                        SFaccountForm.Show vbModal
                        
                        Dim inpt As String
                        inpt = SFaccountForm.TextBox1
                        AccId = "X"                                     ' если связывать не будем - останется так
                        j = 0                                           ' на случай некорректного ввода, к след. account'у
                        If IsNumeric(inpt) Then
                            j = CInt(inpt)
                            If j > 0 And j <= compNum Then
                                AccId = Sheets(SFacc).Cells(CompSNums(j), SFACC_IDACC_COL)  ' Salesforce id
                                Fruitful = Fruitful + 1
                                GoTo endloop
                            End If
                        ElseIf inpt = "exit" Or inpt = "cont" Then
                            GoTo endloop                            '
                        End If
                        If MsgBox("Некорректное значение номера: '" + inpt + "' Продолжать?", vbYesNo) <> vbYes Then Exit Do
                    Loop
endloop:
                    If inpt = "exit" Then
                        ExRespond = False
                        Exit For
                    End If
                End If
             
                .Cells(i, CSIT_MS_IDSF_COL) = AccId
                
            End If
            
        End With
    Next i
BreakForI:
    ModEnd CSIT_MS
    MsgBox "Найдено " & Fruitful & " (" & Format(Fruitful / (i - 7), "Percent") & ") лидов в SF"
End Sub
'Sub AccntSFAnlz()
'
'    Dim i As Long, j As Long, k As Long
'
'    Const Doing = "Анализ справочника клиентов 1С"
'    ModStart Acc1C, Doing
'
'    CheckSheet Acc1C, 1, 5, "Название фирмы"
'
''---------- Подготовка хеш-таблиц --------------------------------------
''   1. Таблица SFacc (имена целиком)
'
'    Dim accSFComps(0 To 9999) As String, accSFCIndxs(0 To 9999) As String
'
''   2. Таблица клиентов 1С (имена целиком)
'    Dim acc1CNames(0 To 4999) As String, acc1CNIndxs(0 To 4999) As String
'    Lines = EOL(Acc1C) - ACC1C_RES          ' кол-во строк 1С
'    hashInit acc1CNames, acc1CNIndxs
'    For i = 2 To Lines
'        hashSet acc1CNames, acc1CNIndxs, LCase$(RemIgnored(Sheets(Acc1C).Cells(i, A1C_NAME_COL))), ""
'    Next i
''   2. Таблица клиентов 1С (по компонентам)
'    Dim acc1CComps(0 To 9999) As String, acc1CIndxs(0 To 9999) As String
'    Dim x() As String, s0 As String                     ' локальные переменные
'
'    hashInit acc1CComps, acc1CIndxs
'    For i = 2 To Lines
''If InStr(LCase$(RemIgnored(Trim$(Sheets(Acc1C).Cells(i, A1C_NAME_COL)))), "судостр") <> 0 Then
''i = i
''End If
'        x = split(LCase$(RemIgnored(Trim$(Sheets(Acc1C).Cells(i, A1C_NAME_COL)))))
'        For j = 0 To UBound(x)
'            s0 = hashGet(acc1CComps, acc1CIndxs, x(j))      ' допустимо несколько
'            If s0 <> "$" Then                               ' индексов для одного account'a
'                s0 = s0 + "$"                               ' разделенных "$"
'            Else
'                s0 = ""
'            End If
'            hashSet acc1CComps, acc1CIndxs, x(j), s0 + str(i)
'        Next j
'    Next i
''---------- Проход по справочнику предприятий SF --------------------------
'
'    EOL_SFacc = EOL(SFacc) - SFresLines
'
'    Fruitful = 0
'
''    ProgressForm.Show vbModeless
''    ProgressForm.ProgressLabel.Caption = Doing
'
'' ищем пред SF по словам 1С
'
'    Dim SFname As String, wr1C As String
'    Dim sfWrds() As String, SFWordIndx As Long
'    Dim MSG As String, MSG2 As String
'    Dim CompSNums(1 To 100) As Long, compNum As Long                ' номерa компонент имени
'    Dim nam1C(1 To 100) As String, count(1 To 100) As Long     '
'    Dim adrTxt(1 To 100) As String, kword(1 To 100) As String
'    Dim adrField As String, AdrSruct As PostAddr
'    Dim DlgRes As String
'
'    For i = 2 To EOL_SFacc
'        If ExRespond = False Then GoTo ExitSub
'        With Sheets(SFacc)
'            If .Cells(i, SFACC_ACC1C_COL) = "" Then
'                ' 1С имя не заполнено. попытаемся найти
'                SFname = .Cells(i, SFACC_ACCNAME_COL)
'                MSG = "ИМЯ Salesforce:     " + SFname + vbCrLf _
'                    + "АДРЕС:              " + SFPostAddr(i, SFacc)
'                MSG2 = ""
'                compNum = 0
'                ' Разбиваем на слова и формируем запрос
'                ' ищем слова в хеше слов 1C
'                sfWrds = split(LCase$(RemIgnored(SFname)))
'If i = 55 Then
'i = i
'End If
'                For SFWordIndx = 0 To UBound(sfWrds)
'                    wr1C = hashGet(acc1CComps, acc1CIndxs, sfWrds(SFWordIndx))
'                    If wr1C <> "$" Then
'                        x = split(wr1C, "$")
'                        For j = 0 To UBound(x)
'                            adrField = Sheets(Acc1C).Cells(x(j), A1C_ADR_COL)
'                            If adrField <> "" Then    ' пропускаем строки без адреса
''                                AdrSruct = AdrParse(adrField)
'                                compNum = compNum + 1           ' считаем варианты
'                                CompSNums(compNum) = i          ' запоминаем номер строки в SFacc
'                                nam1C(compNum) = Sheets(Acc1C).Cells(x(j), A1C_NAME_COL)
'                                adrTxt(compNum) = adrField
'
'
'                                kword(compNum) = sfWrds(SFWordIndx)
'                           End If
'                        Next j
'                    End If
'                Next SFWordIndx
'
'If compNum > 0 Then
'i = i
'End If
'
'' дедупликация и сортировка по количеству совпадений.
'' чем больше совпадений, тем выше приоритет и тем ближе к началу.
'
'
'                For j = 1 To compNum        ' иницализация счетчиков дупликатов
'                    count(j) = 1
'                Next j
'                For j = 1 To compNum - 1       ' поиск дупликатов
'                    For k = j + 1 To compNum
'                        If k <> j Then
'                            If nam1C(k) = nam1C(j) Then
'                                count(j) = count(j) + 1     ' нашли дупликат
'                                count(k) = -1000            ' пометить на удаление
'                                kword(j) = kword(j) + " " + kword(k)
'                            End If
'                        End If
'                    Next k
'                Next j
'
'                ' сортировка по числу дупликатов
'                For j = 1 To compNum
'                    For k = j + 1 To compNum
'                        If count(k) > count(j) Then
'                            switch count, j, k
'                            switch nam1C, j, k
'                            switch CompSNums, j, k
'                            switch adrTxt, j, k
'                            switch kword, j, k
'                        End If
'                    Next k
'                Next j
'
'                ' финальная поготовка текста
'                For j = 1 To compNum
'                    If count(j) > 0 Then
'                        MSG2 = MSG2 + vbCrLf + vbCrLf + "            " _
'                            + Format(j) + ".     '" + "'" + nam1C(j) + "'            КЛЮЧИ: " + kword(j) _
'                            + vbCrLf & "                 Адрес: " + adrTxt(j)
'                    End If
'                Next j
'
'' Текст подготовлен. Запускаем диалог.
'                DlgRes = DlgAccChoice(Acc1C, CompSNums, SFACC_ACCNAME_COL, MSG, MSG2)
'                If IsNumeric(DlgRes) Then
'                    MsgBox "выбрано " + DlgRes
'                End If
'            End If
'       End With
'    Next i
'ExitSub:
'End Sub
Sub Client1CAnlz()

    Dim i As Long, j As Long, k As Long
    Dim x() As String, s0 As String                     ' локальные переменные
    
    ModStart Acc1C, "Анализ справочника клиентов 1С"
    
    CheckSheet SFacc, EOL_SFacc + 2, 1, SFaccRepName
    CheckSheet Acc1C, 1, 5, "Название фирмы"
    
    ClearSheet A_Acc, Range("HDR_AdAcc")
    ClearSheet AccntUpd, Range("HDR_AccntUpd")
    
'---------- Подготовка хеш-таблиц --------------------------------------
'   1. Таблица SFacc (ключ - SF имена целиком, значение - индекс в таблице SFacc)

    Dim accSF(0 To 9999) As String, accSFind(0 To 9999) As String
    hashInit accSF, accSFind
    For i = 2 To EOL_SFacc
        hashSet accSF, accSFind, Compressor(Sheets(SFacc).Cells(i, SFACC_ACC1C_COL)), CInt(i)
    Next i

'   2. Таблица SFacc (ключ - слово из имени предприятия SF,
'                     значение - номера в SFacc, разделеленные символом '$'

    Dim accSFComps(0 To 9999) As String, accSFIndxs(0 To 9999) As String
    
    hashInit accSFComps, accSFIndxs
    For i = 2 To EOL_SFacc
        x = split(LCase$(RemIgnored(Trim$(Sheets(SFacc).Cells(i, SFACC_ACCNAME_COL)))))
        For j = 0 To UBound(x)
            s0 = hashGet(accSFComps, accSFIndxs, x(j))      ' допустимо несколько
            If s0 <> "$" Then                               ' индексов для одного account'a
                s0 = s0 + "$"                               ' разделенных "$"
            Else
                s0 = ""
            End If
            hashSet accSFComps, accSFIndxs, x(j), s0 + Str(i)
        Next j
    Next i
'---------- Проход по справочнику клиентов 1С --------------------------

' EOL_AccntUpd - строка выходного файла - модификация предприятий SF
' EOL_AdAcc    - строка выходного файла - создание предприятий SF
    
'    ProgressForm.Show vbModeless
'    ProgressForm.ProgressLabel.Caption = Doing

' ищем пред SF по словам 1С

    Dim accntName As String, accntNamePrev As String, wrSF As String
    Dim accntAddr As String
    Dim sfWrds() As String, SFWordIndx As Long
    Dim Msg As String, MSG2 As String
    Dim CompSNums(1 To 100) As Long, compNum As Long                ' номерa компонент имени
    Dim namSF(1 To 100) As String, count(1 To 100) As Long          ' SFacc полное имя
    Dim adrTxt(1 To 100) As String, kword(1 To 100) As String
    Dim adrField As String, AdrSruct As PostAddr
    Dim DlgRes As String

    For i = 2 To EOL_Acc1C                                          ' цикл по справочнику 1С
        If ExRespond = False Then GoTo ExitSub
        
        With Sheets(Acc1C)
            accntAddr = .Cells(i, A1C_ADR_COL)
                    
            If accntAddr <> "" Then ' если поле 'адрес' (почтовый адрес) не заполнено, пропускаем
            
                accntName = .Cells(i, A1C_NAME_COL)
                ' дедупликация по 1С имени. считаем, что отсортировано по имени предприятия
                If accntName <> accntNamePrev Then
                   
                    ' проверяем, нет ли уже ссылки из какого-нибудь SF account'a на это предприятие 1С
                    If hashGet(accSF, accSFind, Compressor(accntName)) = "$" Then
                        
                        Msg = Str(i) + ":  ИМЯ 1С:     " + accntName + vbCrLf _
                            + "АДРЕС:              " + .Cells(i, A1C_ADR_COL)
                        MSG2 = ""
                        compNum = 0
                        ' Разбиваем 1С имя на слова и формируем запрос: ищем слова в хеше слов SF
                        sfWrds = split(LCase$(RemIgnored(accntName)))
                        For SFWordIndx = 0 To UBound(sfWrds)
                            wrSF = hashGet(accSFComps, accSFIndxs, sfWrds(SFWordIndx))
                            If wrSF <> "$" Then
                                x = split(wrSF, "$")
                                For j = 0 To UBound(x)
                                    adrField = SFPostAddr(x(j), SFacc)
                                    If adrField <> "" Then    ' пропускаем строки без адреса
                                        compNum = compNum + 1           ' считаем варианты
                                        CompSNums(compNum) = x(j)       ' запоминаем номер строки SFacc
                                        namSF(compNum) = Sheets(SFacc).Cells(x(j), SFACC_ACCNAME_COL)
                                        adrTxt(compNum) = adrField
                                        
                                        kword(compNum) = sfWrds(SFWordIndx)
                                   End If
                                Next j
                            End If
                        Next SFWordIndx
                        
        ' namSF - массив имен в SF, adrTxt - массив соответствуюших адресов,
        ' CompSNums - соответствующие номера строк в SFacc
        ' compNum - число элементов
        
        ' дедупликация и сортировка по количеству совпадений.
        ' чем больше совпадений, тем выше приоритет и тем ближе к началу.
        
                        For j = 1 To compNum        ' иницализация счетчиков дупликатов
                            count(j) = 1
                        Next j
                        For j = 1 To compNum - 1       ' поиск дупликатов
                            For k = j + 1 To compNum
                                If k <> j Then
                                    If namSF(k) = namSF(j) Then
                                        count(j) = count(j) + 1     ' нашли дупликат
                                        count(k) = -1000            ' пометить дупликат на удаление
                                        kword(j) = kword(j) + " " + kword(k)
                                    End If
                                End If
                            Next k
                        Next j
                        
'                   сортировка по числу дупликатов (пузырьком).
'                   Сами дупликаты с count = -1000 помещаются в конец и не включаются в текст
                        For j = 1 To compNum
                            For k = j + 1 To compNum
                                If count(k) > count(j) Then
                                    switch count, j, k
                                    switch namSF, j, k
                                    switch CompSNums, j, k
                                    switch adrTxt, j, k
                                    switch kword, j, k
                                End If
                            Next k
                        Next j
        
'                   финальная подготовка текста
                        For j = 1 To compNum
                            If count(j) > 0 Then
                                MSG2 = MSG2 + vbCrLf + vbCrLf + "            " _
                                    + Format(j) + ".     '" + "'" + namSF(j) + "'            КЛЮЧИ: " + kword(j) _
                                    + vbCrLf & "                 Адрес: " + adrTxt(j)
                            Else
                                compNum = j - 1             ' модифицируем ограничение для диалога
                                GoTo endLoopPrepTxt         ' и форсируем выход из цикла
                            End If
                        Next j
endLoopPrepTxt:
        
'                   Текст подготовлен. Запускаем диалог.
                        DlgRes = DlgAccChoice(CompSNums, compNum, A1C_NAME_COL, Msg, MSG2, namSF, adrTxt, kword)
                        If IsNumeric(DlgRes) Then  ' SF account id  + 1C id
'                            MsgBox "выбрано " + DlgRes + vbCrLf + " (" _
'                                + Sheets(Acc1C).Cells(i, A1C_NAME_COL) + "; " _
'                                + Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_ACCNAME_COL) + "')" _

                            EOL_AccntUpd = EOL_AccntUpd + 1
                        ' связка: пара(id SF предприятия, 1С имя)
                            With Sheets(AccntUpd)
                                .Cells(EOL_AccntUpd, ACCUPD_SFID_COL) = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_IDACC_COL)
                                .Cells(EOL_AccntUpd, ACCUPD_1CNAME_COL) = accntName     ' имя из справочника 1С -> SF account
                            End With
                        ElseIf DlgRes = "create" Then
                            ' заполняем поля формы: имя 1С, имя SF (сейчас они идентичны)
                            NewSFaccForm.Adr1C.Caption = Sheets(Acc1C).Cells(i, A1C_NAME_COL)
                            NewSFaccForm.SFacc.value = NewSFaccForm.Adr1C.Caption
                            
                            ' почтовый адрес
                            s0 = Trim(Sheets(Acc1C).Cells(i, A1C_ADR_COL))
                            AdrSruct = AdrParse(s0)
                            NewSFaccForm.Area.value = AdrSruct.State
                            NewSFaccForm.City.value = AdrSruct.City
                            NewSFaccForm.Street.value = AdrSruct.Street
                            NewSFaccForm.Index.value = AdrSruct.PostIndex
                            NewSFaccForm.Country.value = AdrSruct.Country
                            ' фактический адрес(1С) / адрес доставки(SF)
                            s0 = Trim(Sheets(Acc1C).Cells(i, A1C_FACTADR_COL))
                            AdrSruct = AdrParse(s0)
                            NewSFaccForm.AreaD.value = AdrSruct.State
                            NewSFaccForm.CityD.value = AdrSruct.City
                            NewSFaccForm.StreetD.value = AdrSruct.Street
                            NewSFaccForm.IndexD.value = AdrSruct.PostIndex
                            NewSFaccForm.CountryD.value = AdrSruct.Country
                            
                            NewSFaccForm.contact.value = Sheets(Acc1C).Cells(i, A1C_CON_COL)
                            Dim INN
                            INN = Sheets(Acc1C).Cells(i, A1C_INN_COL)
                            If INN <> "" Then INN = split(INN, "/")(0)
                            NewSFaccForm.INN = INN
                            NewSFaccForm.phone.value = Sheets(Acc1C).Cells(i, A1C_TEL_COL)
                            
                            ' справочные поля - не вводятся
                            NewSFaccForm.invoice.Caption = Sheets(Acc1C).Cells(i, A1C_INVOICE_COL)
                            NewSFaccForm.good.Caption = Sheets(Acc1C).Cells(i, A1C_GOOD_COL)
                            
                            NewSFaccForm.Show vbModal
                            
                            DlgRes = NewSFaccForm.result.value
                            If DlgRes = "save" Then
                                EOL_AdAcc = EOL_AdAcc + 1
'                                MsgBox NewSFaccForm.SFacc.value _
'                                    + vbCrLf + NewSFaccForm.Adr1C.value _
'                                    + vbCrLf + NewSFaccForm.City.value _
'                                    + vbCrLf + NewSFaccForm.Area.value _
'                                    + vbCrLf + NewSFaccForm.Street.value _
'                                    + vbCrLf + NewSFaccForm.Index.value _
'                                    + vbCrLf + NewSFaccForm.Country.value
                            With Sheets(A_Acc)
                                .Cells(EOL_AdAcc, ADACC_NAME_COL) = NewSFaccForm.SFacc
                                .Cells(EOL_AdAcc, ADACC_1CNAME_COL) = NewSFaccForm.Adr1C
                                .Cells(EOL_AdAcc, ADACC_CITY_COL) = NewSFaccForm.City.value
                                .Cells(EOL_AdAcc, ADACC_STATE_COL) = NewSFaccForm.Area.value
                                .Cells(EOL_AdAcc, ADACC_STREET_COL) = NewSFaccForm.Street.value
                                .Cells(EOL_AdAcc, ADACC_INDEX_COL) = NewSFaccForm.Index.value
                                .Cells(EOL_AdAcc, ADACC_COUNTRY_COL) = NewSFaccForm.Country.value
                                .Cells(EOL_AdAcc, ADACC_CONTACT1C_COL) = NewSFaccForm.contact.value
                                .Cells(EOL_AdAcc, ADACC_INN_COL) = NewSFaccForm.INN
                                .Cells(EOL_AdAcc, ADACC_TEL_COL) = NewSFaccForm.phone.value
                                .Cells(EOL_AdAcc, ADACC_FACTCITY_COL) = NewSFaccForm.CityD.value
                                .Cells(EOL_AdAcc, ADACC_FACTSTATE_COL) = NewSFaccForm.AreaD.value
                                .Cells(EOL_AdAcc, ADACC_FACTSTREET_COL) = NewSFaccForm.StreetD.value
                                .Cells(EOL_AdAcc, ADACC_FACTINDEX_COL) = NewSFaccForm.IndexD.value
                                .Cells(EOL_AdAcc, ADACC_FACTCOUNTRY_COL) = NewSFaccForm.CountryD.value
                            End With
                            
                            End If
                            
                        End If
                    End If              ' конец if по ссылкам в SF
                End If                  ' конец if по дедупликации
            End If                      ' конец if по полю '1С адрес
        End With
    Next i
ExitSub:
    
    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\Account\"
    WriteCSV A_Acc, "AdAcc.txt"
    Shell "quota_Acc.bat"
    
    WriteCSV AccntUpd, "AccntUpd.txt"
    Shell "quotaAccUpd.bat"
    
    MS "ИТОГ: created " + Str(EOL_AdAcc) + " linked " + Str(EOL_AccntUpd)
End Sub

Function SFPostAddr(ByVal indx As Long, SFacc As String)
'   Стандартное представление почтового адреса
' 31.05.12

    With Sheets(SFacc)
        SFPostAddr = Replace((.Cells(indx, SFACC_CITY_COL) _
                + "," + .Cells(indx, SFACC_STATE_COL) _
                + "," + .Cells(indx, SFACC_STREET_COL) _
                + "," + .Cells(indx, SFACC_INDEX_COL) _
                + "," + .Cells(indx, SFACC_COUNTRY_COL)), ",,", ",")
    End With
End Function
Function DlgAccChoice(CompSNums, count, idCol, Msg, MSG2, namSF, addrTxt, kword)
    ' CompSNums - массив номеров строк в таблице
    ' count     - actual possibility count
    ' idCol     - номер колонки в таблице
    ' MSG       - заголовок запроса к оператору, часть, не зависящая от выбора
    ' MSG2      - список предложений для выбора
    
    Dim i As Long
    
    If count = 0 Then
        DlgAccChoice = "create"       ' the only possibility
        Exit Function
    End If
    
    DlgAccChoice = "cont"       ' если связывать не будем - останется так
 '   If MSG2 <> "" Then
        SFaccountForm.accntChoice.ColumnCount = 3
        Do
            ' заполнение listbox
            Do While SFaccountForm.accntChoice.ListCount <> 0
                SFaccountForm.accntChoice.RemoveItem 0
            Loop
            For i = 1 To count
                SFaccountForm.accntChoice.AddItem
                SFaccountForm.accntChoice.List(i - 1, 0) = namSF(i)
                SFaccountForm.accntChoice.List(i - 1, 1) = addrTxt(i)
                SFaccountForm.accntChoice.List(i - 1, 2) = kword(i)
            Next i
                        
            SFaccountForm.TextBox2.value = Msg          ' + MSG2       ' основной текст
            If count = 1 Then
                SFaccountForm.TextBox1.value = "1"
            Else
                SFaccountForm.accntChoice.ListIndex = -1
                SFaccountForm.TextBox1.value = ""               ' исходное значение номера - пусто
            End If
            
            ' ok button & textbox enabled only when count <> 0
'            SFaccountForm.TextBox1.Visible = True
            SFaccountForm.OKButton.Visible = True
            If count = 0 Then
'                SFaccountForm.TextBox1.Visible = False
                SFaccountForm.OKButton.Visible = False
            End If
            SFaccountForm.Show vbModal
            
            Dim inpt As String, j As Long
            inpt = SFaccountForm.TextBox1
            j = 0                                           ' на случай некорректного ввода, к след. account'у
            If IsNumeric(inpt) Then
                j = CInt(inpt)
                If j > 0 And j <= count Then
                    DlgAccChoice = j
                    GoTo endloop
                End If
            ElseIf inpt = "exit" Or inpt = "cont" Or inpt = "create" Then
                DlgAccChoice = inpt
                GoTo endloop                            '
            End If
            If MsgBox("Некорректное значение номера: '" + inpt + "' Продолжать?", vbYesNo) <> vbYes Then Exit Do
        Loop
endloop:
        If inpt = "exit" Then ExRespond = False
'    End If

End Function
Function switch(kword, j, k)

' 2 компоненты массива меняются местами
' 5.6.2012
    Dim s As String
    s = kword(j)
    kword(j) = kword(k)
    kword(k) = s
End Function


