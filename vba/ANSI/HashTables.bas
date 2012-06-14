Attribute VB_Name = "HashTables"
'---------------------------------------------------------------------------------
' HashTables  - Работа с hash - таблицами
'
' Работа с hash - таблицами (сейчас только aDicKey - слова, входящие в названия предприятий)
'
' T testfillADic()          - test fillADic
' - fillADic(keyTable, valTable) - заполнение hash-таблицы по листу A_Dic
' T testFindAccHash()       - test findAccHash
' - findAccHash(Client)     - извлечь значение из hash-таблицы A_Dic
'------------------------------------------------
' - пакет Hash: hashInit, hashSet, hashGet, hashFnc
' 24.05.12
Option Explicit

Dim DicReadFlag As Boolean              ' статически инициализируется в False

Dim aDicKey(0 To 9999) As String        ' hash словарь - ключи
Dim aDicVal(0 To 9999) As String        ' значения
Sub testfillADic()
    fillADic aDicKey, aDicVal
End Sub
Sub fillADic(aDicKey, aDicVal)
' заполнение hash-таблицы по листу A_Dic (слова употребляемые в именах предприятий)
' 24.05.12

' ключ - 1-я колонка, значение - 2-я

    Dim i As Integer, x As String
    
    hashInit aDicKey, aDicVal
    For i = 2 To 9999
        x = Sheets(A_Dic).Cells(i, 1)                               ' ключ
        If x = "" Then
            GoTo ExitSub                                            ' конец таблицы
        End If
        hashSet aDicKey, aDicVal, x, Sheets(A_Dic).Cells(i, 2)      ' вставляем в таблицу пару ключ/значение
    Next i
ExitSub:
End Sub

Sub testFindAccHash()
    Dim a(0 To 10) As String
    a(1) = findAccHash("xxxxxx")
    a(2) = findAccHash("чегевара")
End Sub
Function findAccHash(Client)
' извлечь значение из hash-таблицы A_Dic
' 24.05.12

    If Not DicReadFlag Then        ' предполагаем, что DicReadFlag статически инициализирован как False
'        hashInit accHTbl, accVTbl
        fillADic aDicKey, aDicVal
'        fillSFacc
'        fillAcc1C
        DicReadFlag = True
    End If
    
    Dim accWords() As String, accName As String
    
    Dim i As Long
    
    findAccHash = ""
    accWords = split(Client, " ")
    For i = LBound(accWords) To UBound(accWords)
        accName = hashGet(aDicKey, aDicVal, Trim$(accWords(i)))
        If accName <> "$" Then
            findAccHash = accName
            GoTo ExitSub            ' Goto замещает End Sub
        End If
    Next i
ExitSub:
End Function

'------------------------------------------------
' Поддержка hash-таблицы А.Пасс 3.5.12

Sub hashTest()
' Тестовая программа

    Const HASH_SIZE = 100

    Dim hashKey(0 To HASH_SIZE - 1) As String
    Dim hashVal(0 To HASH_SIZE - 1) As String
    
    hashInit hashKey, hashVal

    If hashGet(hashKey, hashVal, "xxxx") <> "$" Then MsgBox " error - нашли ""xxxx"", хотя не заносили"

    hashSet hashKey, hashVal, "xxxx", ""
    hashSet hashKey, hashVal, "yyyy", "1"

    If hashGet(hashKey, hashVal, "xxxx") <> "" Then MsgBox " error - нашли " _
        & hashGet(hashKey, hashVal, "xxxx") & " для ""хххх"""
    If hashGet(hashKey, hashVal, "yyyy") <> "1" Then MsgBox " error - нашли " _
        & hashGet(hashKey, hashVal, "yyyy") & " для ""yyyy"""
End Sub


'-----------------------------------------------
' Инициализация таблицы

Sub hashInit(hashKey, hashVal)
    Dim i As Integer
    Dim size As Integer
    size = UBound(hashKey) - LBound(hashKey) + 1
    
    For i = 0 To size - 1
        hashKey(i) = "$"
        hashVal(i) = "$"     ' вернем "$", если не нашли
    Next i
End Sub

Function hashFnc(ByVal key As String, ByVal size As Integer)
'
' 17.05.12
' псевдослучайная функция, значения которой равномерно и плотно
' заполняют диапазон целых чисел от 0 до size-1

    Dim lng As Integer
    Dim rslt As Long
    Dim i As Integer
    
    If key = "" Then
        hashFnc = 0     ' hashFnc("") = 0
        Exit Function
    End If
    
    lng = Len(key)
    rslt = 0
    
    If lng > 20 Then
        ' берем 20 равнмерно выбранных символов строки (каждый n-й, где n ~ lng / 20)
        For i = 1 To 20
            rslt = ((rslt * 99) Mod size) + Asc(Mid(key, i * lng \ 20, 1))
        Next i
    Else
        ' или все символы строки, если строка короткая
        For i = 1 To lng
            rslt = ((rslt * 99) Mod size) + Asc(Mid(key, i, 1))
        Next i
    End If

    hashFnc = (rslt * 77) Mod size
    
End Function
Sub hashSet(hashKey, hashVal, ByVal key As String, ByVal value As String)
'
' сохранить пару key-value в hash-таблице
'   3.5.12 А.Пасс

    Dim h, h0 As Integer
    Dim size As Integer
    
    size = UBound(hashKey) - LBound(hashKey) + 1
    
    h0 = hashFnc(key, size)
    h = h0
    Do
        If hashKey(h) = "$" Or hashKey(h) = key Then    ' свободная ячейка или key найден
            hashVal(h) = value
            hashKey(h) = key
            Exit Sub
        Else
            h = (h + 1) Mod size
            If h = h0 Then
                MsgBox "Таблица переполнена. Аварийная остановка hashSet", , "FATAL ERROR"
                Stop
            End If
        End If
    Loop
    Stop        ' сюда попасть не можем, выход из цикла - только exit sub
    
End Sub

Function hashGet(hashKey, hashVal, ByVal key As String)

' извлечь value по значению key из hash-таблицы.
' если не найдено, возвращается значение '$'

    Dim h, h0 As Integer
    Dim size As Integer
    Dim hkey As String
    
    size = UBound(hashKey) - LBound(hashKey) + 1

    h0 = hashFnc(key, size)
    h = h0
    Do
        hkey = hashKey(h)
        If hkey = key Or hkey = "$" Then
            hashGet = hashVal(h)        ' поиск окончен; '$' - если не нашли
            Exit Function
        Else
            h = (h + 1) Mod size
            If h = h0 Then
                hashGet = "$"           ' не нашли, таблица полна
                Exit Function
            End If
        End If
    Loop

End Function
