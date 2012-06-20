Attribute VB_Name = "AccountAnalitics"
'---------------------------------------------------------------------------------
' AccountAnalitics  - анализ Имен Организаций и слияние в внешними базами данных
'
' [*] SFaccDicBuild     - проход по листу SFacc для заполнения Словаря Организаций
' (*) PaymentAccPass() - Проход по листу Платежей 1С для внесения
'                        новых Организаций в SF
'  -  RemIgnored(Client) - удаление игнорируемых слов из строки Client
'''''  -  FindAcc(Client)   - поиск Организации в SF по Словарю A_Dic
' x?x NewAcc(Client)    - занесение новой Организации в SF
'  ?  Adr1c(Client)     - адрес клиента 1С или CSIS_MS
'  ?  AdrSF(id)         - адрес Организации по ее Id в SF
'
'   19.6.2012

Option Explicit

Sub SFaccDicBuild()
'
' Проход по листу SFacc для построения Словаря Организаций в SF на листе A_Dic
'   19.4.12
'  7.5.12 - во второй колонке словаря имя организации в 1С
' 23.5.12 - используем SheetDedup2 для А_Dic
' 19.6.12 - добавил колонку 3 IdSFass в A_DIC

    Dim i, j As Integer
    
    Dim Acc As String
    Dim Acc1C As String
    Dim IdSFacc As String
    Dim accWords() As String
    
    Lines = ModStart(SFacc, "SFaccDicBuild: Построение Словаря Организаций", True) - SFresLines
    
    CheckSheet SFacc, EOL_SFacc + 2, 1, SFaccRepName
    ClearSheet A_Dic, Range("HDR_AccDic")

    For i = 2 To Lines
        Progress (i / Lines)
        If ExRespond = False Then Exit For
        
        Acc = LCase$(Sheets(SFacc).Cells(i, SFACC_ACCNAME_COL))
        Acc1C = Sheets(SFacc).Cells(i, SFACC_ACC1C_COL)
        IdSFacc = Sheets(SFacc).Cells(i, SFACC_IDACC_COL)
        accWords = split(RemIgnored(Acc), " ")
        
        For j = LBound(accWords) To UBound(accWords)
            EOL_DIC = EOL_DIC + 1
            Sheets(A_Dic).Cells(EOL_DIC, 1) = accWords(j)
            Sheets(A_Dic).Cells(EOL_DIC, 2) = Acc1C
            Sheets(A_Dic).Cells(EOL_DIC, 3) = IdSFacc
        Next j
    Next i

    Call SheetDedup2(A_Dic, 1, 2, 3)
    
    ModEnd A_Dic
End Sub
Sub PaymentAccPass()
'
' (*) PaymentAccPass() - Проход по листу Платежей 1С для внесения
'                        новых Организаций в SF
'   19.4.12

    Dim i, j As Integer
    
    Dim Acc As String
    Dim accWords() As String
    
    Lines = ModStart(1, "Занесение новых Организаций из Платежей 1С", True) - 3
    
    CheckSheet 1, 1, 6, Stamp1Cpay1    ' проверяем правильность отчета по платежам 1С
    CheckSheet 1, 1, 7, Stamp1Cpay2
    ClearSheet A_Acc, Range("HDR_AdAcc")    ' готовим лист новых Организаций
    EOL_DIC = EOL(A_Dic)
    EOL_Acc1C = EOL(Acc1C)
    
    For i = 2 To Lines
        Progress (i / Lines)
        If ExRespond = False Then Exit For
        
If i Mod 50 = 0 Then
i = i
End If
        If Sheets(1).Cells(i, 1) <> 1 Then
            Acc = Sheets(1).Cells(i, PAYACC_COL)
            If FindAcc(Acc) = "$" Then
                AdAccFr1C Acc           ' такой Организации нет в SF - занесем!
            End If
        End If
    Next i

    Call SheetDedup(A_Acc, 1)   ' возможно несколько платежей по новой организации
    
    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\Account\"
    WriteCSV A_Acc, "AdAcc.txt"
    Shell "quota_Acc.bat"
    
    ModEnd A_Dic
End Sub
Sub testRemIgnored()
    Dim cmp As Integer
    Dim t As String
    
    t = RemIgnored("OOO")
    cmp = StrComp("ооо", "аорг", vbTextCompare)
End Sub

Function RemIgnored(Client) As String
'
'Удаляем все символы пунктуации, аббревиатуры и односимвольные слова
'   20.4.12
'   20.5.12 - перезагружена из SF
'   21.5.12 - отладка

    Dim i As Long, j As Long, rLow As Long, rUp As Long
    Dim s As String

    s = RemDelimiters(LCase$(Client))
    
    Dim Swords() As String
    Dim pattern As String
    Dim NotFound As Boolean
    
    Dim tword As String
    Dim low As Integer
    Dim up As Integer
    Dim cmp As Integer
    
    RemIgnored = ""
    Swords = split(Compressor(s), " ")
      
    rLow = Sheets(We).Range("Glossary").Row - 1
    rUp = Sheets(We).Range("Glossary").End(xlDown).Row + 1
    
    For i = LBound(Swords) To UBound(Swords)    ' цикл по символам Client
        tword = Swords(i)
        If Len(tword) > 1 Then      ' однобуквенные слова игнорируем
            '           up & low проверке не подлежат
            
            ' двоичный поиск tword в Range("Glossary")
            
            low = rLow
            up = rUp
            
            NotFound = True
            Do While NotFound And (up - low) > 1
                j = (low + up) \ 2                  ' округляем с недостатком
                
                ' принимается: j <> up and j <> low;
                
                pattern = LCase$(Sheets(We).Cells(j, 1))   ' берем из листа 'We'
                cmp = StrComp(tword, pattern, vbTextCompare)
                If cmp = 0 Then
                    NotFound = False                ' совпало
                ElseIf Sheets(We).Cells(j, 2) <> "" _
                        And Left$(tword, Len(pattern)) = pattern Then
                    NotFound = False      ' не совпало, но подходит (суффиксы игнорируем)
                ElseIf cmp > 0 Then
                    low = j                         ' идем вверх по 'We'
                Else                                ' ergo, pattern < tword
                    up = j                          ' идем вниз по 'We'
                End If
            Loop
            If NotFound Then
                If RemIgnored <> "" Then RemIgnored = RemIgnored + " "
                RemIgnored = RemIgnored + tword
            End If
        End If      ' (Len > 1)
    Next i
    
End Function
Function NewAcc(Client) As String
'
' формирование новой Организации в SF по данным Справочника 1С
'   3/4/12

If InStr(Client, "Гипробум") > 0 Then
Client = Client
End If

    Dim SeekWords() As String
    Dim SFwords() As String
    Dim SFname As String
    Dim i, j, k
    Dim Msg, Respond
    Dim AccId, Id As String
    
    NewAcc = "": AccId = ""
' ----------- убираем все аббревиатуры ------------------------
    SFname = RemIgnored(LCase$(Client))
    SeekWords = split(SFname, " ")
' ----- ищем совпадающие с именем Client слова в SFacc ----------

    With Sheets(SFacc)
        For i = 2 To EOL_SFacc
            SFname = .Cells(i, SFACC_ACCNAME_COL)
'            On Error Resume Next
'            id = ""
'            id = WorksheetFunction.VLookup(SFname, _
'                Sheets(SFacc).Range("A:C"), 3, False)
'            On Error GoTo 0
'            If id = "" Then
'                MsgBox "FATAL ERROR NewAcc - сбой Id"
'                Stop
'            End If

            SFwords = split(RemDelimiters(LCase$(SFname)), " ")
            For j = 0 To UBound(SeekWords)
                For k = 0 To UBound(SFwords)
                    If SeekWords(j) = SFwords(k) And Len(SeekWords(j)) > 1 Then
                        AccId = "X"
'                        MSG = "<><> 1C = '" & Client & "';" _
'                            & vbCrLf & "Адрес: " & Adr1c(Client) _
'                            & vbCrLf & vbCrLf & " В SF '" & SFname & "'" _
'                            & vbCrLf & "Адрес: " & AdrSF(id) _
'                            & vbCrLf & vbCrLf & "Совпало слово '" & SFwords(K) & "'"
'                        Respond = MsgBox(MSG, vbYesNoCancel)
'                        If Respond = vbYes Then
'                            Fruitful = Fruitful + 1
'                        ' дедупликация Id: если такая связка уже есть - NOP
'                            If InStr(AccId, id) = 0 Then
'                                If AccId <> "" Then AccId = AccId & ","
'                                AccId = AccId & id
'                                LogWr "<><> NewAcc: '" & SFname & "' = '" _
'                                    & Client & "' " & AccId
'                            End If
'                        ElseIf Respond = vbCancel Then
'                            ExRespond = False
'                            NewAcc = AccId
'                            Exit Function
'                        End If
                    End If
                Next k
            Next j
        Next i
    End With
    If AccId = "" Then AccId = "*"
    NewAcc = AccId
End Function
Function Adr1C(Client) As String
'
' Adr1c(Client) - адрес клиента 1С или CSIS_MS - пока пустышка
'
End Function
Function AdrSF(Id)
'
' AdrSF(id) - адрес Организации по ее Id в SF
'
End Function
Sub testIsAccSF()
    Dim i(10), A(10)
'    A(1) = IsAccSF("abbyy", i(1))
'    A(2) = IsAccSF("Акрон", i(2))
'    A(3) = IsAccSF("Роспроект", i(3))
    A(4) = IsAccSF("ООО «Газпром трансгаз Уфа»", i(4))
    A(2) = IsAccSF("Гипротюменьнефтегаз", i(2))
    A(3) = IsAccSF("Студия-44", i(3))
End Sub

Function IsAccSF(Str, iSFacc) As Boolean
'
' - IsAccSF(Str, iSFacc) - возвращает TRUE и номер строки в SFacc,
'                      если найдена Огранизация в SF по строке Str
'   19.6.12

    Dim s() As String   '= строка Str разбитая на слова по пробелам
    Dim AccId As String '= Id Организации из A_Dic
    Dim Id0 As String   '= Id по первому слову в Str
    Dim i As Integer, j As Integer
    Dim x() As String   '= массив Id по слову в Str
    
    IsAccSF = False: iSFacc = 0
    If Str = "" Then Exit Function
    
    s = split(RemIgnored(Str), " ")
    
    For i = LBound(s) To UBound(s)
        AccId = ""
        On Error Resume Next
        AccId = WorksheetFunction.VLookup(s(i), _
            Sheets(A_Dic).Range("A:C"), 3, False)
        On Error GoTo 0
        
        If i = LBound(s) Then
            Id0 = AccId
        Else
            Id0 = xId(Id0, AccId)
        End If
    Next i
    If Id0 = "" Then Exit Function
    If InStr(Id0, "+") <> 0 Then
        ErrMsg TYPE_ERR, "Необходим диалог для выбора Организации"
    Stop
    End If

    iSFacc = WorksheetFunction.Match(Id0, _
        Sheets(SFacc).Range("C:C"), 0)
    IsAccSF = True
End Function
Function xId(Id1, Id2) As String
'
' - xId(Id1, Id2) - пересечение строк Id1 и Id2 разделенных "+"
'   20.6.12

    Dim s() As String
    Dim i As Integer
    
    s = split(Id2, "+")
    xId = ""
    For i = LBound(s) To UBound(s)
        If InStr(Id1, s(i)) <> 0 Then
            If xId <> "" Then xId = xId & "+"
            xId = xId & s(i)
        End If
    Next i
End Function
