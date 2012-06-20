Attribute VB_Name = "AccountAnalitics"
'---------------------------------------------------------------------------------
' AccountAnalitics  - ������ ���� ����������� � ������� � �������� ������ ������
'
' [*] SFaccDicBuild     - ������ �� ����� SFacc ��� ���������� ������� �����������
' (*) PaymentAccPass() - ������ �� ����� �������� 1� ��� ��������
'                        ����� ����������� � SF
'  -  RemIgnored(Client) - �������� ������������ ���� �� ������ Client
'''''  -  FindAcc(Client)   - ����� ����������� � SF �� ������� A_Dic
' x?x NewAcc(Client)    - ��������� ����� ����������� � SF
'  ?  Adr1c(Client)     - ����� ������� 1� ��� CSIS_MS
'  ?  AdrSF(id)         - ����� ����������� �� �� Id � SF
'
'   19.6.2012

Option Explicit

Sub SFaccDicBuild()
'
' ������ �� ����� SFacc ��� ���������� ������� ����������� � SF �� ����� A_Dic
'   19.4.12
'  7.5.12 - �� ������ ������� ������� ��� ����������� � 1�
' 23.5.12 - ���������� SheetDedup2 ��� �_Dic
' 19.6.12 - ������� ������� 3 IdSFass � A_DIC

    Dim i, j As Integer
    
    Dim Acc As String
    Dim Acc1C As String
    Dim IdSFacc As String
    Dim accWords() As String
    
    Lines = ModStart(SFacc, "SFaccDicBuild: ���������� ������� �����������", True) - SFresLines
    
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
            If FindAcc(Acc) = "$" Then
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
    Dim t As String
    
    t = RemIgnored("OOO")
    cmp = StrComp("���", "����", vbTextCompare)
End Sub

Function RemIgnored(Client) As String
'
'������� ��� ������� ����������, ������������ � �������������� �����
'   20.4.12
'   20.5.12 - ������������� �� SF
'   21.5.12 - �������

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
    Dim Msg, Respond
    Dim AccId, Id As String
    
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
Function Adr1C(Client) As String
'
' Adr1c(Client) - ����� ������� 1� ��� CSIS_MS - ���� ��������
'
End Function
Function AdrSF(Id)
'
' AdrSF(id) - ����� ����������� �� �� Id � SF
'
End Function
Sub testIsAccSF()
    Dim i(10), A(10)
'    A(1) = IsAccSF("abbyy", i(1))
'    A(2) = IsAccSF("�����", i(2))
'    A(3) = IsAccSF("���������", i(3))
    A(4) = IsAccSF("��� �������� �������� ���", i(4))
    A(2) = IsAccSF("�������������������", i(2))
    A(3) = IsAccSF("������-44", i(3))
End Sub

Function IsAccSF(Str, iSFacc) As Boolean
'
' - IsAccSF(Str, iSFacc) - ���������� TRUE � ����� ������ � SFacc,
'                      ���� ������� ����������� � SF �� ������ Str
'   19.6.12

    Dim s() As String   '= ������ Str �������� �� ����� �� ��������
    Dim AccId As String '= Id ����������� �� A_Dic
    Dim Id0 As String   '= Id �� ������� ����� � Str
    Dim i As Integer, j As Integer
    Dim x() As String   '= ������ Id �� ����� � Str
    
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
        ErrMsg TYPE_ERR, "��������� ������ ��� ������ �����������"
    Stop
    End If

    iSFacc = WorksheetFunction.Match(Id0, _
        Sheets(SFacc).Range("C:C"), 0)
    IsAccSF = True
End Function
Function xId(Id1, Id2) As String
'
' - xId(Id1, Id2) - ����������� ����� Id1 � Id2 ����������� "+"
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
