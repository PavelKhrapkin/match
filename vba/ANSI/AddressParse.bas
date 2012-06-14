Attribute VB_Name = "AddressParse"
'---------------------------------------------------------------------------------
' AddressParse  - ������ ������� ����������� � ������� � �������� ������ ������
'
' T TestAddr()        - ������� AdrParse - ������ ������ ������
' - AdrParse(Address) - ������ ������ ��������� ������
' - SeekDIC - ��������� ���������� �������� �� ������� � Range(DicRange) ��� Adr
' - adrCompRemove   - ������� ���������� ������.
' - AdAccFr1C(Acc)  - ���������� ����� ����������� Acc �� 1� � ���� A_Acc ��� ��������� � SF
' T testfillADic    - fillADic
' - fillADic()      -  ���������� hash-������� �� ����� A_Dic (����� ������������� � ������ �����������)
' T testfillSFacc   - ���� fillSFacc
' - fillSFacc       - ���������� hash-������� �� ����� SFacc
' T testfillAcc1C
' - fillAcc1C       - ��������� hash �� Acc1C - ������� ����������� 1�
'   20.5.2012   A.����
    
Option Explicit

' hash ������� ��� AdAccFr1C - ����� ��������
Dim accHTbl(0 To 5000) As String
Dim accVTbl(0 To 5000) As String

Dim DicReadFlag As Boolean              ' ���������� ���������������� ��� False
Dim aDicKey(0 To 9999) As String        ' hash �������
Dim aDicVal(0 To 9999) As String

Dim SFaccKey(0 To 4999) As String       ' hash ������� ����������� � 1�
Dim SFaccVal(0 To 4999) As String
Dim Acc1CKey(0 To 4999) As String       ' hash ������� ����������� � SF
Dim Acc1CVal(0 To 4999) As String
 
Sub TestAddr()
'
' T testAddr() - ������� AdrParse - ������ ������ ������
'   14.5.12 �.����

    Dim A1(0 To 20) As PostAddr

    A1(18) = AdrParse("196105. �. �����-���������X.��������� ��., �.11     ")
    A1(16) = AdrParse(" �. �����-���������. ��. ��������� �. 1-3, ���. / ����  592-43-60")
    A1(17) = AdrParse("198323, ������������� �������, ������������� �����, � ����� �������, ����� ����������, 7")
    A1(15) = AdrParse(" ���������� ���. ������������� � �.  ����� ������� ��. ��������, �.1 143026")
    
    ' ��������� � ��������
    
    A1(1) = AdrParse("198323, ������������� �������, ������������� �����, � ����� �������, ����� ����������, 7")
    A1(2) = AdrParse("195009, �. �����-���������, ��. ��������� �. 1-3, ���. / ����  592-43-60")
    A1(3) = AdrParse("128234, ������, ��������� 25 �� 8, ������")
    A1(4) = AdrParse("128234, ������, ��������� 25 �� 8, �-��")
    A1(5) = AdrParse("128234, ������, ��������� 25 �� 8, �������")
    A1(6) = AdrParse("AAA")                                                                                         ' ERR
    A1(7) = AdrParse("")                                                                                            ' ERR
    A1(8) = AdrParse("143026, ���������� ���., ����������� �-��,�.�.��������������, ��. ��������, �.1         ")
    A1(9) = AdrParse("143026 ���������� ���., ����������� �-��,�.�.��������������, ��. ��������, �.1         ")
    A1(10) = AdrParse(" 143026, ���������� ���., ����������� �-��,�.�.��������������, ��. ��������, �.1         ")
    A1(11) = AdrParse(" ���������� ���., ����������� �-��,�.�.��������������, ��. ��������, �.1 , 143026")
    A1(12) = AdrParse("    +143026")                                                                                ' ERR

    ' ��������� ��� �������
    
    A1(13) = AdrParse(" ���������� ���. ����������� �-�� �.�.�������������� ��. ��������, �.1 143026")
    A1(14) = AdrParse(" ���������� ���. ����������� �-�� �.�. �������������� ��. ��������, �.1 143026")
    A1(15) = AdrParse(" ���������� ���. ������������� � �.  ����� ������� ��. ��������, �.1 143026")
    
    Dim clIndx As Integer
    Dim addrToParse As String
    
    ' ��������� �� ���� ������ 1�
'
    EOL_AdAcc = 1
    For clIndx = 2 To 9999
        If Sheets(Acc1C).Cells(clIndx, A1C_NAME_COL) = "" Then GoTo endCLoop   ' ������ ��� - ����� �����
        addrToParse = Sheets(Acc1C).Cells(clIndx, A1C_ADR_COL)
        If addrToParse <> "" Then
                
            clIndx = clIndx                     ' ��� �������
            If A1(0).ErrFlag Then
                clIndx = clIndx                     ' ��� �������
            Else
            
                EOL_AdAcc = EOL_AdAcc + 1
'If EOL_AdAcc = 9 Then
'            ' ��������� ������� ��� �������
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
' AdrParse(Address)   - ��������� ������ ��������� ������ Address
'   11.5.12

'    AdrParse = AdrParse(Address, True)
''    If AdrParse.ErrFlag Then
''        AdrParse = AdrParse(Address, False)    ' ���� �� ���������� - ������� ��� �������
''    End If
'End Function
'
'
'Function AdrParse(ByVal Address As String, commaSplit As Boolean) As PostAddr

' - AdrParse(Address) - ������ ������ ��������� ������ � ��������� ����� split
'       (�� �������, ���� true, �� ��������, ���� false)
'   11.5.12

    Dim Adr() As String
    Dim i As Integer, j As Integer
    Dim lastIndxPos As Integer
    
'    Const addrExtraWrds = "� ���������� ���� ��� �� � "
    
    AdrParse.ErrFlag = False

' --- ���� �������� ������

    Address = Address + ","             ' ����� ��������� �������� �� ����� � ����� ������
    AdrParse.PostIndex = ""
    lastIndxPos = -100                  ' ����� �������� ����� �������� �� ���������
    For i = 1 To Len(Address)
        If Not IsNumeric(Mid(Address, i, 1)) Then
            If i - lastIndxPos = 6 Then ' ����� ������������������ ����. ��������� �����
                AdrParse.PostIndex = Mid(Address, lastIndxPos, 6)
                Mid(Address, lastIndxPos, 6) = "      "
                GoTo split              ' ������ ������
            End If
            lastIndxPos = -100
        Else
            If lastIndxPos < 0 Then lastIndxPos = i     ' �����. ��� ������ � �����?
        End If
    Next i
    
split:
    Adr = split(Compressor(Address), ",")
    
'--- ���� �����, �������, ������
    AdrParse.City = SeekDIC("DIC_CityNames", Adr)
    AdrParse.Country = SeekDIC("DIC_CtryNames", Adr)
    AdrParse.State = SeekDIC("DIC_Area", Adr)
    
    If AdrParse.Country = "" Then AdrParse.Country = "������"

'--- ��������� ���� �����
    Dim Street As String
    Dim x As String, curr As String
    
    ' �������� ����������� ���������� � Street ����� �������
    Street = ""
    For i = LBound(Adr) To UBound(Adr)
' ���������� ������ ������ � �������� �����
        x = Trim$(Adr(i))
        For j = 1 To Len(x)
            curr = Mid(x, j, 1)
            ' ���� �������� �������, ����� - �������
            If curr <> " " And curr <> "." Then GoTo breakL     ' �������� break
            Mid(x, j, 1) = " "                                   ' ������� ���������� ����� � ������ �������
        Next j
breakL:
        If x <> "" Then
            If Street <> "" Then Street = Street & ", "
            Street = Street & x
        End If
    Next i
    
    AdrParse.Street = Compressor(Street)
    
'--- ��������� ErrFlag (state ����� ��������)
    If AdrParse.City = "" Or InStr(AdrParse.City, "?") <> 0 _
            Or AdrParse.PostIndex = "" Or AdrParse.Street = "" Then
        AdrParse.ErrFlag = True
    End If
    
End Function
Function SeekDIC(DicRange, Adr) As String
'
' SeekDIC - ��������� ���������� �������� �� ������� � Range(DicRange) ��� Adr
'   14.5.12


    Dim sWe As Range
    Dim wrds() As String
    Dim wrdsArea() As String
    Dim SeekWord As String, pattern As String, PatternArea As String, curr As String, SeekArea As String
    Dim i As Integer, j As Integer, k As Integer, kArea As Integer
    Dim wrdPos As Integer
    
    SeekDIC = ""
    For i = LBound(Adr) To UBound(Adr)      ' ���� �� ����������� ������
        If Adr(i) <> "" Then
            SeekWord = LCase$(Adr(i))
            For Each sWe In Range(DicRange).Rows    ' ���� �� Range (���� ��: DIC_CityNames, DIC_CtryNames, DIC_Area)
                wrds = split(sWe.Cells(1, 1), ",")
                For k = LBound(wrds) To UBound(wrds)    ' ���� �� ������ ���. 1 (���������� ���������, ������)
                    pattern = Trim(LCase$(wrds(k)))     ' Trim ��������� ������ ��� ������� ������ �������
                    wrdPos = InStr(SeekWord, pattern)
                    If wrdPos <> 0 Then
                        
                        ' ��������� ���������� ������� �� �������
                        If DicRange = "DIC_CityNames" Then
                            adrCompRemove Adr(i), wrdPos, pattern
                        Else
                            Adr(i) = ""
                        End If
                        SeekDIC = wrds(LBound(wrds))    ' ������ ������� �������
                        
                        ' ���� �����, ���� ��������� � ���. 2 ����. "We"
                        ' ����� ����������� � ��������� ����� �������, ����� �������

                        If sWe.Cells(1, 2) <> "" Then
                            For j = LBound(Adr) To UBound(Adr)     ' ���� �� ����������� ������
                                If Adr(j) <> "" Then
                                    SeekArea = LCase$(Adr(j))
                                    wrdsArea = split(sWe.Cells(1, 2), ",")
                                    For kArea = LBound(wrdsArea) To UBound(wrdsArea) ' ���� �� ������ ���. 2 (���������� ���������, ������)
                                        PatternArea = Trim(LCase$(wrdsArea(kArea)))
                                        wrdPos = InStr(SeekArea, LCase$(wrdsArea(kArea)))
                                        If wrdPos <> 0 Then
                                            adrCompRemove Adr(j), wrdPos, PatternArea
                                            SeekDIC = wrdsArea(LBound(wrdsArea)) & ", " & SeekDIC   ' ������ ������� �������
                                            Exit Function
                                        End If
                                    Next kArea
                                End If
                            Next j
                            
                            SeekDIC = wrdsArea(LBound(wrdsArea)) & ", ?"        ' ����� ������ �������, ���� �������� � "We"
                            
                        End If
                            
                        Exit Function
                    End If
                Next k
            Next sWe
        End If
    Next i
    

End Function
Sub adrCompRemove(AdrComp, wrdPos, pattern)

'   adrCompRemove - ������� ���������� ������.
'       �������� ������� ������ ����� ���������
'   14.05.2012

    Dim patternEnd As Integer, i As Integer, lastpos As Integer
    Dim curr As String, wspflag As Boolean, wrd As String
    
'    If commaSplit Then
'        AdrComp = ""
'    Else
        ' �������� ��������� ����� ����� ����� ������ �� ������ ��� Len(Pattern)
        patternEnd = wrdPos + Len(pattern) - 1  ' ��������� ������� �����
        lastpos = 0                             ' �� ������, ���� �� ��������
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
       
' ����� �������������� ����� �, ���� ��� ����������� (����., "�"), ������� ���
        If wrdPos > 1 Then
            wspflag = True
            i = wrdPos - 1
            Do
                curr = Mid(AdrComp, i, 1)
                If wspflag Then
                    ' ���������� ������� � ����� ������
                    If curr <> " " Or curr <> "." Then wspflag = False
                Else
                    ' ���� ������������� ������
                    If curr = " " Then
                        i = i + 1       ' �����, ������� ������ �� 1 ������� � �������
                        GoTo remPrefix
                    End If
                End If
                If i = 1 Then
                    GoTo remPrefix      ' ������ ������, �������
                Else
                    i = i - 1           ' ��������� ������ �����
                End If
            Loop
remPrefix:
' ����������� ��������� ����� � Street.
' ����� ���������/��� �������, �������� �������, ������ ����� � ��������� - ���������?
            wrd = Replace(LCase$(Trim(Mid(AdrComp, i, wrdPos - i))), ".", "") & " "
' �����, ����������� � ��������� � Street, ������ ���������. ������ ������ ���� ����� ������� �����,
' ��� ������ ���� � ������ �������� � ��� �����.
            If InStr("� ��� ���������� ���� � ��� �� � ���", wrd) <> 0 Then
                Mid(AdrComp, i, wrdPos - i) = "                                        "
            End If
        End If
        
'    End If
End Sub
Sub testAddAcc1C()

    ClearSheet A_Acc, Range("HDR_AdAcc")
    EOL_Acc1C = EOL(Acc1C) - ACC1C_RES
    
    hashInit accHTbl, accVTbl
    
    AdAccFr1C ("��� ""���������� ��������""")
    AdAccFr1C ("��� ""���������� ��������""")       ' �������� �� ��������
    AdAccFr1C ("������ ��������     ")

End Sub
Sub testAdAccFr1C()

    AdAccFr1C "��� ""�������������"""

    AdAccFr1C "������ ���� �� ��� ������ �� �.�����-���������� � ������������� �������"
    AdAccFr1C "xxxxxxxxxxx"     ' ��� � 1�
    AdAccFr1C "��������"        ' ������ �����

End Sub
Sub iniAdAccFr1C()
'   ������ ��������������� � AdAccFr1C
'    hashInit accHTbl, accVTbl
End Sub

Sub AdAccFr1C(acc)
'
' ���������� ����� ����������� Acc �� 1� � ���� A_Acc ��� ��������� � SF
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
    
    If Not DicReadFlag Then        ' ������������, ��� DicReadFlag ���������� ��������������� ��� False
        hashInit accHTbl, accVTbl
        fillADic
        fillSFacc
        fillAcc1C
        DicReadFlag = True
    End If
    
' ���� ����� ����������� ��� ���� � ����� A_Acc - ������ �� ������
    If hashGet(accHTbl, accVTbl, acc) <> "$" Then GoTo ExitSub
    
' ���� ������ ����� �� account � A_Dic. ���� �������, �� �������, �� ����� � Log
    accWords = split(acc, " ")
    For i = LBound(accWords) To UBound(accWords)
        If hashGet(aDicKey, aDicVal, Trim$(accWords(i))) <> "$" Then
            LogWr "<!> Account '" & acc & "'" _
                & "' �������� ����� '" & accWords(i) & "', ��������� � �������."
            GoTo ExitSub
        End If
    Next i
        
' ����: ���� �� ��� � SF
    If hashGet(SFaccKey, SFaccVal, acc) <> "$" Then
        LogWr "<!> Account '" & acc & "'" & "' ������������ � ���� '��� ����������� � 1�' "
    End If
                
' ������������ - �������� ����������� �� �������
    For i = 2 To EOL_AdAcc
        If acc = Sheets(A_Acc).Cells(i, ADACC_NAME_COL) Then GoTo ExitSub
    Next i
    
' ���� account � ������� 1�, ����� ������� ����� � ��. ������ (��. fillAcc1C)

    accIndxStr = hashGet(Acc1CKey, Acc1CVal, acc)
    If accIndxStr = "$" Then GoTo ExitSub               ' �� ����� � ������� 1�
    accIndx = Val(accIndxStr)                           ' �������� �����
    
    With Sheets(Acc1C)
       
        INN = .Cells(accIndx, A1C_INN_COL)
        If INN <> "" Then INN = split(INN, "/")(0)
            
''''''''''''''''''''''''''''''''''
        addrString = .Cells(accIndx, A1C_ADR_COL)
        If addrString <> "" Then
            Addr = AdrParse(addrString)
            If Addr.ErrFlag Then
                LogWr "<!> ������ ������� ������ ��� '" & acc & "'" _
                    & "; ����� '" & Trim(addrString) & "'"
            Else
                GoTo FoundAdr
            End If
  ''''''''''''''''''''''''''''''''''
        End If
    End With
    GoTo ExitSub            ' ��� � 1� - �������
    
    
            '        For i = 2 To EOL_Acc1C
            '    For i = 1 To EOL_SFacc
            '        ' ��������� �� 1� �����
            '        If acc = Sheets(SFacc).Cells(i, SFACC_ACC1C_COL) Then
            '            LogWr "<!> Account '" & acc & "'" & "' ������������ � ���� '��� ����������� � 1�' "
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
'        .Cells(EOL_AdAcc, ADACC_CONT_COL) = Sheets(Acc1C).Cells(accIndx, A1C_CONT_COL)  ' �������
    End With

' ��������� ��������� � hash ������� (accHTbl,accVTbl)
    hashSet accHTbl, accVTbl, acc, ""
    
ExitSub:
End Sub
Sub testfillADic()
    fillADic
End Sub
Sub fillADic()
' ���������� hash-������� �� ����� A_Dic (����� ������������� � ������ �����������)
' 18.05.12

' ������������ ������ ����, value �� �����������

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
' � ���� fillSFacc
'   19.5.2012
    fillSFacc
End Sub
Sub fillSFacc()
' ���������� hash-������� �� ����� SFacc
' 18.05.12

' ������������ ������ ����, value �� �����������

    Dim i As Long, x As String
    
    hashInit SFaccKey, SFaccVal
    
    Dim ef As Long
    ef = EOL_Acc1C
    If ef = 0 Then ef = 9999                    ' ������ ��� �������

    For i = 2 To ef
        If Sheets(SFacc).Cells(i, 1) <> "" Then                    ' ���������� ������
            x = Sheets(SFacc).Cells(i, SFACC_ACC1C_COL)
            If i = 3589 Then
                i = i
            End If
            If ef = 9999 Then                       ' ������ ��� �������
                If Sheets(SFacc).Cells(i, 1) = "SFacc" Then GoTo ExitSub    ' ������ ��� �������
            End If                                  ' ������ ��� �������
        End If
    Next i
ExitSub:
End Sub
Sub testfillAcc1C()
' � ���� fillAcc1C
'   19.5.2012
    fillSFacc
End Sub
Sub fillAcc1C()

' ��������� hash �� Acc1C - ������� ����������� 1�
'   19.5.2012

    Dim i As Integer, x As String
    
    Dim ef As Long
    ef = EOL_SFacc
    If ef = 0 Then ef = 9999        ' ������ ��� �������

    hashInit Acc1CKey, Acc1CVal
    For i = 1 To ef
        x = Sheets(Acc1C).Cells(i, A1C_NAME_COL)
        If x = "" Then GoTo ExitSub
        If ef = 9999 Then                       ' ������ ��� �������
            If x = "SFacc" Then GoTo ExitSub    ' ������ ��� �������
        End If                                  ' ������ ��� �������
        ' ���������� � hash �������� ���������
        
If x = "��������" Then
i = i
End If
        
        If hashGet(Acc1CKey, Acc1CVal, x) = "$" Then hashSet Acc1CKey, Acc1CVal, x, i
    Next i
ExitSub:
End Sub
