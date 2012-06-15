Attribute VB_Name = "AccountAnalitics"
'---------------------------------------------------------------------------------
' AccountAnalitics  - ������ ���� ����������� � ������� � �������� ������ ������
'
' [*] SFaccDicBuild     - ������ �� ����� SFacc ��� ���������� ������� �����������
' (*) PaymentAccPass() - ������ �� ����� �������� 1� ��� ��������
'                        ����� ����������� � SF
'  -  RemIgnored(Client) - �������� ������������ ���� �� ������ Client
'  -  FindAcc(Client)   - ����� ����������� � SF �� ������� A_Dic
' x?x NewAcc(Client)    - ��������� ����� ����������� � SF
'  ?  Adr1c(Client)     - ����� ������� 1� ��� CSIS_MS
'  ?  AdrSF(id)         - ����� ����������� �� �� Id � SF
'
'   23.5.2012

Option Explicit

Sub SFaccDicBuild()
'
' ������ �� ����� SFacc ��� ���������� ������� ����������� � SF �� ����� A_Dic
'   19.4.12
'  7.5.12 - �� ������ ������� ������� ��� ����������� � 1�
' 23.5.12 - ���������� SheetDedup2 ��� �_Dic

    Dim i, j As Integer
    
    Dim Acc As String
    Dim Acc1C As String
    Dim accWords() As String
    
    Lines = ModStart(SFacc, "SFaccDicBuild: ���������� ������� �����������", True) - SFresLines
    
    CheckSheet SFacc, Lines + 2, 1, SFaccRepName
    ClearSheet A_Dic, Range("HDR_AccDic")

    For i = 2 To Lines
        Progress (i / Lines)
        If ExRespond = False Then Exit For
        
        Acc = LCase$(Sheets(SFacc).Cells(i, SFACC_ACCNAME_COL))
        Acc1C = Sheets(SFacc).Cells(i, SFACC_ACC1C_COL)
        accWords = split(RemIgnored(Acc), " ")
        
        For j = LBound(accWords) To UBound(accWords)
            EOL_DIC = EOL_DIC + 1
            Sheets(A_Dic).Cells(EOL_DIC, 1) = accWords(j)
            Sheets(A_Dic).Cells(EOL_DIC, 2) = Acc1C
        Next j
    Next i

    Call SheetDedup2(A_Dic, 1, 2)
    
    Sheets(A_Dic).Cells(1, 3) = Date & " " & Time
    ModEnd A_Dic
End Sub
Sub PaymentAccPass()
'
' (*) PaymentAccPass() - ������ �� ����� �������� 1� ��� ��������
'                        ����� ����������� � SF
'   19.4.12

    Dim i, j As Integer
    
    Dim Acc As String
    Dim accWords() As String
    
    Lines = ModStart(1, "��������� ����� ����������� �� �������� 1�", True) - 3
    
    CheckSheet 1, 1, 6, Stamp1Cpay1    ' ��������� ������������ ������ �� �������� 1�
    CheckSheet 1, 1, 7, Stamp1Cpay2
    ClearSheet A_Acc, Range("HDR_AdAcc")    ' ������� ���� ����� �����������
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
            If FindAcc(Acc) = "*" Then
                AdAccFr1C Acc           ' ����� ����������� ��� � SF - �������!
            End If
        End If
    Next i

    Call SheetDedup(A_Acc, 1)   ' �������� ��������� �������� �� ����� �����������
    
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Account\"
    WriteCSV A_Acc, "AdAcc.txt"
    Shell "quota_Acc.bat"
    
    ModEnd A_Dic
End Sub
Sub testRemIgnored()
    Dim cmp As Integer
    Dim T As String
    
    T = RemIgnored("OOO")
    cmp = StrComp("���", "����", vbTextCompare)
End Sub

Function RemIgnored(Client) As String
'
'������� ��� ������� ����������, ������������ � �������������� �����
'   20.4.12
'   20.5.12 - ������������� �� SF
'   21.5.12 - �������

    Dim i As Long, j As Long, rLow As Long, rUp As Long
    Dim S As String

    S = RemDelimiters(LCase$(Client))
    
    Dim Swords() As String
    Dim pattern As String
    Dim NotFound As Boolean
    
    Dim tword As String
    Dim low As Integer
    Dim up As Integer
    Dim cmp As Integer
    
    RemIgnored = ""
    Swords = split(Compressor(S), " ")
      
    rLow = Sheets(We).Range("Glossary").Row - 1
    rUp = Sheets(We).Range("Glossary").End(xlDown).Row + 1
    
    For i = LBound(Swords) To UBound(Swords)    ' ���� �� �������� Client
        tword = Swords(i)
        If Len(tword) > 1 Then      ' ������������� ����� ����������
            '           up & low �������� �� ��������
            
            ' �������� ����� tword � Range("Glossary")
            
            low = rLow
            up = rUp
            
            NotFound = True
            Do While NotFound And (up - low) > 1
                j = (low + up) \ 2                  ' ��������� � �����������
                
                ' �����������: j <> up and j <> low;
                
                pattern = LCase$(Sheets(We).Cells(j, 1))   ' ����� �� ����� 'We'
                cmp = StrComp(tword, pattern, vbTextCompare)
                If cmp = 0 Then
                    NotFound = False                ' �������
                ElseIf Sheets(We).Cells(j, 2) <> "" _
                        And Left$(tword, Len(pattern)) = pattern Then
                    NotFound = False      ' �� �������, �� �������� (�������� ����������)
                ElseIf cmp > 0 Then
                    low = j                         ' ���� ����� �� 'We'
                Else                                ' ergo, pattern < tword
                    up = j                          ' ���� ���� �� 'We'
                End If
            Loop
            If NotFound Then
                If RemIgnored <> "" Then RemIgnored = RemIgnored + " "
                RemIgnored = RemIgnored + tword
            End If
        End If      ' (Len > 1)
    Next i
    
End Function
Function FindAcc(Client) As String
'
' FindAcc(Client) - ����� ����������� � SF �� ������� A_Dic
'   8.5.12

    Dim i As Integer
    
    Dim Acc As String, x As String
    Dim accWords() As String
    
' ----------- ������� ��� ������������ ------------------------
    If Client = "" Then GoTo NotFound
    Acc = RemIgnored(Client)
    If Acc = "" Then GoTo NotFound
    accWords = split(Acc, " ")
' ----- ���� ����������� � ������ Client ����� � A_Dic ----------
    For i = LBound(accWords) To UBound(accWords)
        x = ""
        On Error Resume Next
        x = WorksheetFunction.VLookup(accWords(i), _
                Sheets(A_Dic).Range("A:B"), 2, False)
        On Error GoTo 0
        If x <> "" Then
            FindAcc = x
            Exit Function
        End If
    Next i
NotFound:
    FindAcc = "*"
End Function
Function NewAcc(Client) As String
'
' ������������ ����� ����������� � SF �� ������ ����������� 1�
'   3/4/12

If InStr(Client, "��������") > 0 Then
Client = Client
End If

    Dim SeekWords() As String
    Dim SFwords() As String
    Dim SFname As String
    Dim i, j, k
    Dim MSG, Respond
    Dim AccId, id As String
    
    NewAcc = "": AccId = ""
' ----------- ������� ��� ������������ ------------------------
    SFname = RemIgnored(LCase$(Client))
    SeekWords = split(SFname, " ")
' ----- ���� ����������� � ������ Client ����� � SFacc ----------

    With Sheets(SFacc)
        For i = 2 To EOL_SFacc
            SFname = .Cells(i, SFACC_ACCNAME_COL)
'            On Error Resume Next
'            id = ""
'            id = WorksheetFunction.VLookup(SFname, _
'                Sheets(SFacc).Range("A:C"), 3, False)
'            On Error GoTo 0
'            If id = "" Then
'                MsgBox "FATAL ERROR NewAcc - ���� Id"
'                Stop
'            End If

            SFwords = split(RemDelimiters(LCase$(SFname)), " ")
            For j = 0 To UBound(SeekWords)
                For k = 0 To UBound(SFwords)
                    If SeekWords(j) = SFwords(k) And Len(SeekWords(j)) > 1 Then
                        AccId = "X"
'                        MSG = "<><> 1C = '" & Client & "';" _
'                            & vbCrLf & "�����: " & Adr1c(Client) _
'                            & vbCrLf & vbCrLf & " � SF '" & SFname & "'" _
'                            & vbCrLf & "�����: " & AdrSF(id) _
'                            & vbCrLf & vbCrLf & "������� ����� '" & SFwords(K) & "'"
'                        Respond = MsgBox(MSG, vbYesNoCancel)
'                        If Respond = vbYes Then
'                            Fruitful = Fruitful + 1
'                        ' ������������ Id: ���� ����� ������ ��� ���� - NOP
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
Function Adr1c(Client) As String
'
' Adr1c(Client) - ����� ������� 1� ��� CSIS_MS - ���� ��������
'
End Function
Function AdrSF(id)
'
' AdrSF(id) - ����� ����������� �� �� Id � SF
'
End Function

