Attribute VB_Name = "HashTables"
'---------------------------------------------------------------------------------
' HashTables  - ������ � hash - ���������
'
' ������ � hash - ��������� (������ ������ aDicKey - �����, �������� � �������� �����������)
'
' T testfillADic()          - test fillADic
' - fillADic(keyTable, valTable) - ���������� hash-������� �� ����� A_Dic
' T testFindAccHash()       - test findAccHash
' - findAccHash(Client)     - ������� �������� �� hash-������� A_Dic
'------------------------------------------------
' - ����� Hash: hashInit, hashSet, hashGet, hashFnc
' 24.05.12
Option Explicit

Dim DicReadFlag As Boolean              ' ���������� ���������������� � False

Dim aDicKey(0 To 9999) As String        ' hash ������� - �����
Dim aDicVal(0 To 9999) As String        ' ��������
Sub testfillADic()
    fillADic aDicKey, aDicVal
End Sub
Sub fillADic(aDicKey, aDicVal)
' ���������� hash-������� �� ����� A_Dic (����� ������������� � ������ �����������)
' 24.05.12

' ���� - 1-� �������, �������� - 2-�

    Dim i As Integer, x As String
    
    hashInit aDicKey, aDicVal
    For i = 2 To 9999
        x = Sheets(A_Dic).Cells(i, 1)                               ' ����
        If x = "" Then
            GoTo ExitSub                                            ' ����� �������
        End If
        hashSet aDicKey, aDicVal, x, Sheets(A_Dic).Cells(i, 2)      ' ��������� � ������� ���� ����/��������
    Next i
ExitSub:
End Sub

Sub testFindAccHash()
    Dim a(0 To 10) As String
    a(1) = findAccHash("xxxxxx")
    a(2) = findAccHash("��������")
End Sub
Function findAccHash(Client)
' ������� �������� �� hash-������� A_Dic
' 24.05.12

    If Not DicReadFlag Then        ' ������������, ��� DicReadFlag ���������� ��������������� ��� False
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
            GoTo ExitSub            ' Goto �������� End Sub
        End If
    Next i
ExitSub:
End Function

'------------------------------------------------
' ��������� hash-������� �.���� 3.5.12

Sub hashTest()
' �������� ���������

    Const HASH_SIZE = 100

    Dim hashKey(0 To HASH_SIZE - 1) As String
    Dim hashVal(0 To HASH_SIZE - 1) As String
    
    hashInit hashKey, hashVal

    If hashGet(hashKey, hashVal, "xxxx") <> "$" Then MsgBox " error - ����� ""xxxx"", ���� �� ��������"

    hashSet hashKey, hashVal, "xxxx", ""
    hashSet hashKey, hashVal, "yyyy", "1"

    If hashGet(hashKey, hashVal, "xxxx") <> "" Then MsgBox " error - ����� " _
        & hashGet(hashKey, hashVal, "xxxx") & " ��� ""����"""
    If hashGet(hashKey, hashVal, "yyyy") <> "1" Then MsgBox " error - ����� " _
        & hashGet(hashKey, hashVal, "yyyy") & " ��� ""yyyy"""
End Sub


'-----------------------------------------------
' ������������� �������

Sub hashInit(hashKey, hashVal)
    Dim i As Integer
    Dim size As Integer
    size = UBound(hashKey) - LBound(hashKey) + 1
    
    For i = 0 To size - 1
        hashKey(i) = "$"
        hashVal(i) = "$"     ' ������ "$", ���� �� �����
    Next i
End Sub

Function hashFnc(ByVal key As String, ByVal size As Integer)
'
' 17.05.12
' ��������������� �������, �������� ������� ���������� � ������
' ��������� �������� ����� ����� �� 0 �� size-1

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
        ' ����� 20 ��������� ��������� �������� ������ (������ n-�, ��� n ~ lng / 20)
        For i = 1 To 20
            rslt = ((rslt * 99) Mod size) + Asc(Mid(key, i * lng \ 20, 1))
        Next i
    Else
        ' ��� ��� ������� ������, ���� ������ ��������
        For i = 1 To lng
            rslt = ((rslt * 99) Mod size) + Asc(Mid(key, i, 1))
        Next i
    End If

    hashFnc = (rslt * 77) Mod size
    
End Function
Sub hashSet(hashKey, hashVal, ByVal key As String, ByVal value As String)
'
' ��������� ���� key-value � hash-�������
'   3.5.12 �.����

    Dim h, h0 As Integer
    Dim size As Integer
    
    size = UBound(hashKey) - LBound(hashKey) + 1
    
    h0 = hashFnc(key, size)
    h = h0
    Do
        If hashKey(h) = "$" Or hashKey(h) = key Then    ' ��������� ������ ��� key ������
            hashVal(h) = value
            hashKey(h) = key
            Exit Sub
        Else
            h = (h + 1) Mod size
            If h = h0 Then
                MsgBox "������� �����������. ��������� ��������� hashSet", , "FATAL ERROR"
                Stop
            End If
        End If
    Loop
    Stop        ' ���� ������� �� �����, ����� �� ����� - ������ exit sub
    
End Sub

Function hashGet(hashKey, hashVal, ByVal key As String)

' ������� value �� �������� key �� hash-�������.
' ���� �� �������, ������������ �������� '$'

    Dim h, h0 As Integer
    Dim size As Integer
    Dim hkey As String
    
    size = UBound(hashKey) - LBound(hashKey) + 1

    h0 = hashFnc(key, size)
    h = h0
    Do
        hkey = hashKey(h)
        If hkey = key Or hkey = "$" Then
            hashGet = hashVal(h)        ' ����� �������; '$' - ���� �� �����
            Exit Function
        Else
            h = (h + 1) Mod size
            If h = h0 Then
                hashGet = "$"           ' �� �����, ������� �����
                Exit Function
            End If
        End If
    Loop

End Function
