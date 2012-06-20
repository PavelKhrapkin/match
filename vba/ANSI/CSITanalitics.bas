Attribute VB_Name = "CSITanalitics"
'----------------------------------------------------------
' ������ ������� ���������� �� CSIT
'   ���� �.     17.6.12
' - CSIT_MS_Clear()                             - ������� ��������� ����� MS_CSIT
' - CSIT_MS_lead()                              - ������ ���� ����������� - ����� Microsoft CSIT
' - Client1CAnlz                                - ������ �� ������� �������� 1�
' - SFPostAddr(indx As Long, SFacc As String)   - ����������� ������������� ��������� ������
' - DlgAccChoice                                - ����� ����� "����� ����������� SF ��� ����������"

Option Explicit

Sub CSIT_MS_Clear()
' ������� ��������� ����� MS_CSIT
'   31.05.12

    Dim i As Long
    EOL_CSIT_MS = EOL(CSIT_MS)
    For i = 7 To EOL_CSIT_MS                ' ���� �� MS ������������
        Sheets(CSIT_MS).Cells(i, CSIT_MS_IDSF_COL) = ""
    Next i

End Sub

Sub CSIT_MS_lead()
'
' ������ ���� ����������� - ����� Microsoft CSIT
'   31.05.12

' ��������� ������ * � ������ �����������, ���� �� ������� ������������ SFacc,
'                  "X", ���� �������, �� �������� �� ������������ �� ������,
'                  ������ ID, ����������� ���������, ���� �������� �� ������������.
' ������������� ������ ������, ���������� "". ��� ���������� � �������� ��������� ������������ CSIT_MS_Clear()

    Const Doing = "������ �������� ����������� - ����� MS_CSIT"
    ModStart CSIT_MS, Doing
    ProgressForm.Show vbModeless
    ProgressForm.ProgressLabel.Caption = Doing
    LogWr ""
    LogWr Doing
    ExRespond = True
    
    EOL_CSIT_MS = EOL(CSIT_MS)
    EOL_SFacc = EOL(SFacc) - SFresLines

    CheckSheet CSIT_MS, 4, 2, CSIT_MS_STAMP
    CheckSheet Acc1C, 1, 5, "�������� �����"
    CheckSheet SFacc, EOL_SFacc + 2, 1, SFaccRepName
'---------- ������ �� CSIT_MS ------------------------
    Dim i As Long, j As Long, k As Long
    Fruitful = 0

    Dim SFname As String, SeekWords() As String, SNumsS() As String, SNums As Long, AccId As String
    
    ' Hash ������� ���������� ����a �� �������� ����������� � ����� ������ � SFacc (���� - ���������� �����, �����)
    Dim SFAccKTbl(0 To 9999) As String, SFAccVTbl(0 To 9999) As String
    hashInit SFAccKTbl, SFAccVTbl    ' ������ ������� ������
    
    Dim curAcc As String, SFAccNums As String, SFId As String
    Dim Msg As String, MSG2 As String, Client As String, Respond As Long
    
    With Sheets(SFacc)
        For i = 2 To EOL_SFacc
            ' ��� ����������� -> id
'            If i = 4 Then
'                i = i
'            End If
            SFname = RemIgnored(LCase$(.Cells(i, SFACC_ACCNAME_COL)))
            If Trim(SFname) = "" Then
                Msg = "������ ������������ ����� � �����: '" & .Cells(i, SFACC_ACCNAME_COL) & "'. "
                LogWr Msg
'                Respond = MsgBox(MSG & vbCrLf & vbCrLf & " ��� �� ����� ��������?", vbYesNo)
'                If Respond <> vbYes Then GoTo NextI
                GoTo NextI      '��������!!!
                SeekWords = split(LCase$(.Cells(i, SFACC_ACCNAME_COL)), "$")        ' �������� ��� ���� �����
            Else
                SeekWords = split(SFname, " ")                                      ' �������� ����� �� �����������, ��� ������������
            End If
            ' ������������� ��� ����������� ��� ����� � ��� ������� ����� ��������� � ������� ����� ������ SFacc
            For j = 0 To UBound(SeekWords)
                curAcc = hashGet(SFAccKTbl, SFAccVTbl, SeekWords(j))
                If curAcc <> "$" Then
                    curAcc = curAcc + " "
                Else
                    curAcc = ""
                End If
                hashSet SFAccKTbl, SFAccVTbl, SeekWords(j), curAcc + Trim(Str(i))     ' ����� ������ � �������
            Next j
NextI:
        Next i
    End With
    
    Dim compNum As Long   ' ���������� ����� ����������� � SF (array, index)
    Dim MSName As String                    ' ��� ����������� � MS
    Dim SFnComps() As String, sfn As Long   ' ���������� ����� ����������� � SF (array, index)
    Dim CompSNums(1 To 100) As Long         ' �����a ��������� �����
    
    ' ���� ���������� �� MS ����� � hash - ������� ��������� SF - ����
    
    For i = 7 To EOL_CSIT_MS                ' ���� �� MS ������������
        Progress (i - 7) / EOL_CSIT_MS
        If ExRespond = False Then GoTo BreakForI
        AccId = "*"                         ' ���� ������ �� ������ - �������� � �������
        With Sheets(CSIT_MS)
            MSName = .Cells(i, CSIT_MS_NAME_COL)
            If MSName <> "" And Trim(.Cells(i, CSIT_MS_IDSF_COL)) = "" Then         ' ���� �� ����� - ��� ������������
            
                ' ��������� ������ MS �����������
                
                Client = .Cells(i, CSIT_MS_NAME_COL)                                ' ��������� ������ - �������� MS account
                Msg = "CSIT_MS ���:" + "'" & Client & "';" _
                    + vbCrLf + "�����: " + .Cells(i, CSIT_MS_ADDR_COL) _
                    + vbCrLf + vbCrLf + "---- ��������� SF ����� ----"
                MSG2 = ""
                SeekWords = split(RemIgnored(LCase$(MSName)), " ")                  ' ��������� MS ��� �� ����������
                compNum = 0
                For j = 0 To UBound(SeekWords)
                
                        ' ������ SF ������� �� hash ����������� �� ���������� MS-����� (����������� - ������)
                    SFAccNums = hashGet(SFAccKTbl, SFAccVTbl, SeekWords(j))
                    If SFAccNums <> "$" Then                                        ' � hash ���-�� ���� (������ � ������� SFacc, ������� SF - accounts)?
                
                        MSG2 = MSG2 + vbCrLf + vbCrLf + "       �������� �����: " + SeekWords(j)
                        SNumsS = split(SFAccNums, " ")
                        For k = 0 To UBound(SNumsS)                                 ' ���� �� ������� SFacc
                            SNums = CInt(SNumsS(k))
                            SFnComps = split(LCase(RemIgnored(Sheets(SFacc).Cells(SNums, SFACC_ACCNAME_COL))), " ")
                            For sfn = 0 To UBound(SFnComps)                         ' ���� �� ����������� SF-�����
                            
                                If SFnComps(sfn) = SeekWords(j) Then
                                    SFname = Sheets(SFacc).Cells(SNums, SFACC_ACCNAME_COL)
                                    compNum = compNum + 1
                                    MSG2 = MSG2 + vbCrLf + vbCrLf + "            " + Format(compNum) + ".     '" + "'" & SFname & "'" _
                                        & vbCrLf & "                 �����: " & SFPostAddr(SNums, SFacc)
                                        
                                    CompSNums(compNum) = SNums      ' ��������� ����� ������ � SFacc
                                    
                                End If
                                                            
                            Next sfn
                        Next k
                    End If
                Next j

' ����� �����������. ��������� ������.

                If MSG2 <> "" Then
                    Do
                        SFaccountForm.TextBox2.value = Msg + MSG2       ' �������� �����
                        SFaccountForm.TextBox1.value = ""               ' �������� �������� ������ - �����
                        SFaccountForm.Show vbModal
                        
                        Dim inpt As String
                        inpt = SFaccountForm.TextBox1
                        AccId = "X"                                     ' ���� ��������� �� ����� - ��������� ���
                        j = 0                                           ' �� ������ ������������� �����, � ����. account'�
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
                        If MsgBox("������������ �������� ������: '" + inpt + "' ����������?", vbYesNo) <> vbYes Then Exit Do
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
    MsgBox "������� " & Fruitful & " (" & Format(Fruitful / (i - 7), "Percent") & ") ����� � SF"
End Sub
'Sub AccntSFAnlz()
'
'    Dim i As Long, j As Long, k As Long
'
'    Const Doing = "������ ����������� �������� 1�"
'    ModStart Acc1C, Doing
'
'    CheckSheet Acc1C, 1, 5, "�������� �����"
'
''---------- ���������� ���-������ --------------------------------------
''   1. ������� SFacc (����� �������)
'
'    Dim accSFComps(0 To 9999) As String, accSFCIndxs(0 To 9999) As String
'
''   2. ������� �������� 1� (����� �������)
'    Dim acc1CNames(0 To 4999) As String, acc1CNIndxs(0 To 4999) As String
'    Lines = EOL(Acc1C) - ACC1C_RES          ' ���-�� ����� 1�
'    hashInit acc1CNames, acc1CNIndxs
'    For i = 2 To Lines
'        hashSet acc1CNames, acc1CNIndxs, LCase$(RemIgnored(Sheets(Acc1C).Cells(i, A1C_NAME_COL))), ""
'    Next i
''   2. ������� �������� 1� (�� �����������)
'    Dim acc1CComps(0 To 9999) As String, acc1CIndxs(0 To 9999) As String
'    Dim x() As String, s0 As String                     ' ��������� ����������
'
'    hashInit acc1CComps, acc1CIndxs
'    For i = 2 To Lines
''If InStr(LCase$(RemIgnored(Trim$(Sheets(Acc1C).Cells(i, A1C_NAME_COL)))), "�������") <> 0 Then
''i = i
''End If
'        x = split(LCase$(RemIgnored(Trim$(Sheets(Acc1C).Cells(i, A1C_NAME_COL)))))
'        For j = 0 To UBound(x)
'            s0 = hashGet(acc1CComps, acc1CIndxs, x(j))      ' ��������� ���������
'            If s0 <> "$" Then                               ' �������� ��� ������ account'a
'                s0 = s0 + "$"                               ' ����������� "$"
'            Else
'                s0 = ""
'            End If
'            hashSet acc1CComps, acc1CIndxs, x(j), s0 + str(i)
'        Next j
'    Next i
''---------- ������ �� ����������� ����������� SF --------------------------
'
'    EOL_SFacc = EOL(SFacc) - SFresLines
'
'    Fruitful = 0
'
''    ProgressForm.Show vbModeless
''    ProgressForm.ProgressLabel.Caption = Doing
'
'' ���� ���� SF �� ������ 1�
'
'    Dim SFname As String, wr1C As String
'    Dim sfWrds() As String, SFWordIndx As Long
'    Dim MSG As String, MSG2 As String
'    Dim CompSNums(1 To 100) As Long, compNum As Long                ' �����a ��������� �����
'    Dim nam1C(1 To 100) As String, count(1 To 100) As Long     '
'    Dim adrTxt(1 To 100) As String, kword(1 To 100) As String
'    Dim adrField As String, AdrSruct As PostAddr
'    Dim DlgRes As String
'
'    For i = 2 To EOL_SFacc
'        If ExRespond = False Then GoTo ExitSub
'        With Sheets(SFacc)
'            If .Cells(i, SFACC_ACC1C_COL) = "" Then
'                ' 1� ��� �� ���������. ���������� �����
'                SFname = .Cells(i, SFACC_ACCNAME_COL)
'                MSG = "��� Salesforce:     " + SFname + vbCrLf _
'                    + "�����:              " + SFPostAddr(i, SFacc)
'                MSG2 = ""
'                compNum = 0
'                ' ��������� �� ����� � ��������� ������
'                ' ���� ����� � ���� ���� 1C
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
'                            If adrField <> "" Then    ' ���������� ������ ��� ������
''                                AdrSruct = AdrParse(adrField)
'                                compNum = compNum + 1           ' ������� ��������
'                                CompSNums(compNum) = i          ' ���������� ����� ������ � SFacc
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
'' ������������ � ���������� �� ���������� ����������.
'' ��� ������ ����������, ��� ���� ��������� � ��� ����� � ������.
'
'
'                For j = 1 To compNum        ' ������������ ��������� ����������
'                    count(j) = 1
'                Next j
'                For j = 1 To compNum - 1       ' ����� ����������
'                    For k = j + 1 To compNum
'                        If k <> j Then
'                            If nam1C(k) = nam1C(j) Then
'                                count(j) = count(j) + 1     ' ����� ��������
'                                count(k) = -1000            ' �������� �� ��������
'                                kword(j) = kword(j) + " " + kword(k)
'                            End If
'                        End If
'                    Next k
'                Next j
'
'                ' ���������� �� ����� ����������
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
'                ' ��������� ��������� ������
'                For j = 1 To compNum
'                    If count(j) > 0 Then
'                        MSG2 = MSG2 + vbCrLf + vbCrLf + "            " _
'                            + Format(j) + ".     '" + "'" + nam1C(j) + "'            �����: " + kword(j) _
'                            + vbCrLf & "                 �����: " + adrTxt(j)
'                    End If
'                Next j
'
'' ����� �����������. ��������� ������.
'                DlgRes = DlgAccChoice(Acc1C, CompSNums, SFACC_ACCNAME_COL, MSG, MSG2)
'                If IsNumeric(DlgRes) Then
'                    MsgBox "������� " + DlgRes
'                End If
'            End If
'       End With
'    Next i
'ExitSub:
'End Sub
Sub PaymntCl1CAnlz()

'   ��������� �������� 1� �� ����� "���� ����� ��������"
'       19.06.12

    Dim i As Long, j As Long, k As Long

    ModStart P_Paid, "������ ���� ����� ��������"
    
    CheckSheet SFacc, EOL_SFacc + 2, 1, SFaccRepName
    CheckSheet PAY_SHEET, 1, PAYDOC_COL, Stamp1Cpay1
    
    ClearSheet A_Acc, Range("HDR_AdAcc")
    ClearSheet AccntUpd, Range("HDR_AccntUpd")
    
    For i = 2 To EOL_Acc1C
        If ExRespond = False Then GoTo ExitSub
        client1CProcess Sheets(PAY_SHEET).Cells(i, PAYACC_COL)      ' �������� - ��� ������� 1�
    Next i
ExitSub:
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Account\"
    WriteCSV A_Acc, "AdAcc.txt"
    Shell "quota_Acc.bat"
    
    WriteCSV AccntUpd, "AccntUpd.txt"
    Shell "quotaAccUpd.bat"
    
    MS "����: created " + Str(EOL_AdAcc) + " linked " + Str(EOL_AccntUpd)
End Sub
Sub Cl1CAnlz()

'   ��������� �������� 1� �� ����� "������ �������� 1C"
'       19.06.12

    Dim i As Long, j As Long, k As Long

    ModStart Acc1C, "������ ����������� �������� 1�"
    
    CheckSheet SFacc, EOL_SFacc + 2, 1, SFaccRepName
    CheckSheet Acc1C, 1, A1C_NAME_COL, ACC1C_STAMP
    
    ClearSheet A_Acc, Range("HDR_AdAcc")
    ClearSheet AccntUpd, Range("HDR_AccntUpd")
    
    For i = 2 To EOL_Acc1C
        If ExRespond = False Then GoTo ExitSub
        If Trim(Sheets(Acc1C).Cells(i, A1C_ADR_COL)) <> "" Then     ' ���������� ������ ��� ������ (����. "�����")
            client1CProcess Sheets(Acc1C).Cells(i, A1C_NAME_COL)    ' �������� - ��� ������� 1�
        End If
    Next i
ExitSub:
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Account\"
    WriteCSV A_Acc, "AdAcc.txt"
    Shell "quota_Acc.bat"
    
    WriteCSV AccntUpd, "AccntUpd.txt"
    Shell "quotaAccUpd.bat"
    
    MS "����: created " + Str(EOL_AdAcc) + " linked " + Str(EOL_AccntUpd)
End Sub

Sub client1CProcess(ByVal accntName As String)

'   ��������� ������ ������� 1�
'   accntName - ��� �������
'       19.06.12

    Dim s0 As String, s1() As String                    ' ��������� ����������
    Dim i As Long, j As Long, k As Long
    
    Static hashFlag As Boolean                              ' ���������������� � False
    Static accSF(0 To 9999) As String, accSFind(0 To 9999) As String
    Static accSFComps(0 To 9999) As String, accSFIndxs(0 To 9999) As String
    Static acc1CHashKey(0 To 4999) As String, acc1CHashVal(0 To 4999) As String
    If (Not hashFlag) Then
    
    '---------- ���������� ���-������ --------------------------------------
    '   1. ������� SFacc (���� - SF ����� �������, �������� - ������ � ������� SFacc)
    
        hashInit accSF, accSFind
        For i = 2 To EOL_SFacc
            hashSet accSF, accSFind, Compressor(Sheets(SFacc).Cells(i, SFACC_ACC1C_COL)), CInt(i)
        Next i
    
    '   2. ������� SFacc (���� - ����� �� ����� ����������� SF,
    '                     �������� - ������ � SFacc, ������������� �������� '$'
    
        
        hashInit accSFComps, accSFIndxs
        For i = 2 To EOL_SFacc
            s1 = split(LCase$(RemIgnored(Trim$(Sheets(SFacc).Cells(i, SFACC_ACCNAME_COL)))))
            For j = 0 To UBound(s1)
                s0 = hashGet(accSFComps, accSFIndxs, s1(j))     ' ��������� ���������
                If s0 <> "$" Then                               ' �������� ��� ������ account'a
                    s0 = s0 + "$"                               ' ����������� "$"
                Else
                    s0 = ""
                End If
                hashSet accSFComps, accSFIndxs, s1(j), s0 + Str(i)
            Next j
        Next i
        hashInit acc1CHashKey, acc1CHashVal                     ' ������������� ���� ������������
        For i = 2 To EOL_Acc1C
        
            If Trim(Sheets(Acc1C).Cells(i, A1C_ADR_COL)) <> "" Then
                hashSet acc1CHashKey, acc1CHashVal, LCase$(Sheets(Acc1C).Cells(i, A1C_NAME_COL)), Str(i)
            End If
        Next i
        hashFlag = True
    End If
    
' ��������� ������� 1�
    Dim clIndx As Long              ' ������ ������� � ������� �������� 1�
    Dim accntPostAddr As String     ' �������� ����� ������� 1�
    Dim Msg As String               ' ��������� �����
    Dim sfWrds() As String, wrSF As String, SFWordIndx As Long
    Dim adrTxt(1 To 100) As String, kword(1 To 100) As String
    Dim adrField As String
    Dim CompSNums(1 To 100) As Long, compNum As Long                ' �����a ��������� �����
    Dim namSF(1 To 100) As String, count(1 To 100) As Long          ' SFacc ������ ���
    
    Dim AdrStruct As PostAddr, AdrStruct1C As PostAddr
    Dim delAddrSF As PostAddr, factAddr1C As PostAddr
    Dim DlgRes As String                                            ' ��������� ������� DlgAccChoice

    s0 = hashGet(acc1CHashKey, acc1CHashVal, LCase$(accntName))
    If s0 = "$" Then
        ErrMsg TYPE_ERR, "client1CProcess: ������� '" + accntName + "' ��� � ����������� �������� 1�"
        GoTo exitProc           ' �������
    End If
    clIndx = CInt(s0)
    If clIndx > 0 Then          ' ������������. ���� �� > 0, ��� ������������.
        With Sheets(Acc1C)
            accntPostAddr = .Cells(clIndx, A1C_ADR_COL)
            If accntPostAddr = "" Then GoTo exitProc   ' ���� ���� '�����' (�������� �����) �� ���������, ����������
            
            ' ���������, ��� �� ��� ������ �� ������-������ SF account'a �� ��� ����������� 1�
            If hashGet(accSF, accSFind, Compressor(accntName)) <> "$" Then GoTo exitProc
            
            ' ������� ��� ������������
            hashSet acc1CHashKey, acc1CHashVal, LCase$(accntName), "-" + s0
            
        ' ������������ ���������� �� 1� � ������ ���������
            Msg = s0 + ":  ��� 1�:     " + accntName + vbCrLf _
                + "�����:              " + .Cells(clIndx, A1C_ADR_COL)
'                        MSG2 = ""
            compNum = 0
            ' ��������� 1� ��� �� ����� � ��������� ������: ���� ����� � ���� ���� SF
            sfWrds = split(LCase$(RemIgnored(accntName)))
            For SFWordIndx = 0 To UBound(sfWrds)
                wrSF = hashGet(accSFComps, accSFIndxs, sfWrds(SFWordIndx))
                If wrSF <> "$" Then
                    s1 = split(wrSF, "$")
                    For j = 0 To UBound(s1)
                        adrField = SFPostAddr(s1(j), SFacc)
                        If adrField <> "" Then    ' ���������� ������ ��� ������
                            compNum = compNum + 1           ' ������� ��������
                            CompSNums(compNum) = s1(j)       ' ���������� ����� ������ SFacc
                            namSF(compNum) = Sheets(SFacc).Cells(s1(j), SFACC_ACCNAME_COL)
                            adrTxt(compNum) = adrField
                            
                            kword(compNum) = sfWrds(SFWordIndx)
                       End If
                    Next j
                End If
            Next SFWordIndx
        End With
                        
        ' namSF - ������ ���� � SF, adrTxt - ������ ��������������� �������,
        ' CompSNums - ��������������� ������ ����� � SFacc
        ' compNum - ����� ���������
        
        ' ������������ � ���������� �� ���������� ����������.
        ' ��� ������ ����������, ��� ���� ��������� � ��� ����� � ������.
    
        For j = 1 To compNum        ' ������������ ��������� ����������
            count(j) = 1
        Next j
        For j = 1 To compNum - 1       ' ����� ����������
            For k = j + 1 To compNum
                If k <> j Then
                    If namSF(k) = namSF(j) Then
                        count(j) = count(j) + 1     ' ����� ��������
                        count(k) = -1000            ' �������� �������� �� ��������
                        kword(j) = kword(j) + " " + kword(k)
                    End If
                End If
            Next k
        Next j
                            
        ' ���������� �� ����� ���������� (���������).
        ' ���� ��������� � count = -1000 ���������� � ����� � �� ���������� � �����
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
            
        ' ��������� ���������� ������
        For j = 1 To compNum
            If count(j) <= 0 Then
                compNum = j - 1             ' ������������ ����������� ��� �������
                GoTo endLoopPrepTxt         ' � ��������� ����� �� �����
            End If
        Next j
endLoopPrepTxt:
    
    
            
        ' ����� �����������. ��������� ������.
        Dim Repeat As Boolean
        Do
            Repeat = False
            DlgRes = DlgAccChoice(CompSNums, compNum, A1C_NAME_COL, Msg, namSF, adrTxt, kword)
            If IsNumeric(DlgRes) Then  ' SF account id  + 1C id
        
            ' ��������� � �������� �����
                SFaccMergeWith1C.SFacc = accntName
                SFaccMergeWith1C.name1C = namSF(CInt(DlgRes))
                SFaccMergeWith1C.setInn Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_INN_COL), _
                                        Sheets(Acc1C).Cells(clIndx, A1C_INN_COL)
                SFaccMergeWith1C.setTel Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_TEL_COL), _
                                        Sheets(Acc1C).Cells(clIndx, A1C_TEL_COL)
        
                ' ���������� �������� ����� �����
                AdrStruct.City = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_CITY_COL)
                AdrStruct.Street = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_STREET_COL)
                AdrStruct.State = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_STATE_COL)
                AdrStruct.PostIndex = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_INDEX_COL)
                AdrStruct.Country = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_COUNTRY_COL)
                AdrStruct1C = AdrParse(Sheets(Acc1C).Cells(clIndx, A1C_ADR_COL))
                delAddrSF.City = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_DELCITY_COL)
                delAddrSF.Street = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_DELSTREET_COL)
                delAddrSF.State = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_DELSTATE_COL)
                delAddrSF.PostIndex = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_DELINDEX_COL)
                delAddrSF.Country = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_DELCOUNTRY_COL)
                factAddr1C = AdrParse(Sheets(Acc1C).Cells(clIndx, A1C_FACTADR_COL))
                SFaccMergeWith1C.setAddr AdrStruct, AdrStruct1C, delAddrSF, factAddr1C
        
                SFaccMergeWith1C.setTel Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_TEL_COL), _
                                        Sheets(Acc1C).Cells(clIndx, A1C_TEL_COL)
        
                SFaccMergeWith1C.Show                               ' ����� �����
        
                If SFaccMergeWith1C.result = "exit" Then            ' ��������� ����������� �����
                    ExRespond = False
                ElseIf SFaccMergeWith1C.result = "save" Then
                    EOL_AccntUpd = EOL_AccntUpd + 1
                    With Sheets(AccntUpd)
                        .Cells(EOL_AccntUpd, ACCUPD_SFID_COL) = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_IDACC_COL)
                        .Cells(EOL_AccntUpd, ACCUPD_1CNAME_COL) = accntName     ' ��� �� ����������� 1� -> SF account
                        .Cells(EOL_AccntUpd, ACCUPD_INN_COL) = SFaccMergeWith1C.innSF
                        .Cells(EOL_AccntUpd, ACCUPD_TEL_COL) = SFaccMergeWith1C.telSF
                    End With
                ElseIf SFaccMergeWith1C.result = "back" Then
                    Repeat = True       ' ������������ ������ ���������� ���������� ����� do
                End If  '���� �� ���� ������� �� ��������� - ������ "����������'
            ElseIf DlgRes = "create" Then
                ' ��������� ���� �����: ��� 1�, ��� SF (������ ��� ���������)
                NewSFaccForm.Adr1C.Caption = Sheets(Acc1C).Cells(clIndx, A1C_NAME_COL)
                NewSFaccForm.SFacc.value = NewSFaccForm.Adr1C.Caption
        
                ' �������� �����
                s0 = Trim(Sheets(Acc1C).Cells(clIndx, A1C_ADR_COL))
                AdrStruct = AdrParse(s0)
                NewSFaccForm.Area.value = AdrStruct.State
                NewSFaccForm.City.value = AdrStruct.City
                NewSFaccForm.Street.value = AdrStruct.Street
                NewSFaccForm.Index.value = AdrStruct.PostIndex
                NewSFaccForm.Country.value = AdrStruct.Country
                ' ����������� �����(1�) / ����� ��������(SF)
                s0 = Trim(Sheets(Acc1C).Cells(clIndx, A1C_FACTADR_COL))
                AdrStruct = AdrParse(s0)
                NewSFaccForm.AreaD.value = AdrStruct.State
                NewSFaccForm.CityD.value = AdrStruct.City
                NewSFaccForm.StreetD.value = AdrStruct.Street
                NewSFaccForm.IndexD.value = AdrStruct.PostIndex
                NewSFaccForm.CountryD.value = AdrStruct.Country
        
                NewSFaccForm.contact.value = Sheets(Acc1C).Cells(clIndx, A1C_CON_COL)
                Dim INN
                INN = Trim(Sheets(Acc1C).Cells(clIndx, A1C_INN_COL))
                If INN <> "" Then INN = split(INN, "/")(0)
                NewSFaccForm.INN = Trim(INN)
                NewSFaccForm.phone.value = Sheets(Acc1C).Cells(clIndx, A1C_TEL_COL)
        
            ' ���������� ���� - �� ��������
                NewSFaccForm.invoice.Caption = Sheets(Acc1C).Cells(clIndx, A1C_INVOICE_COL)
                NewSFaccForm.good.Caption = Sheets(Acc1C).Cells(clIndx, A1C_GOOD_COL)
        
                NewSFaccForm.Show vbModal
        
                DlgRes = NewSFaccForm.result.value
                If DlgRes = "exit" Then
                    ExRespond = False
                ElseIf DlgRes = "save" Then
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
                End If      ' Dlgres= 'exit'
            End If          ' isnumeric(dlgres)
        Loop While Repeat
    End If                  'end if �� ������������
    
exitProc:
End Sub
Sub Client1CAnlz()

' ������ �� ������� �������� 1� � ������� ��������������� �������� SF
' ��� ���������� � ��������� ����������� �������� SF
'   17.6.12

    Dim i As Long, j As Long, k As Long
    Dim x() As String, s0 As String                     ' ��������� ����������
    
    ModStart Acc1C, "������ ����������� �������� 1�"
    
    CheckSheet SFacc, EOL_SFacc + 2, 1, SFaccRepName
    CheckSheet Acc1C, 1, A1C_NAME_COL, ACC1C_STAMP
    
    ClearSheet A_Acc, Range("HDR_AdAcc")
    ClearSheet AccntUpd, Range("HDR_AccntUpd")

'---------- ���������� ���-������ --------------------------------------
'   1. ������� SFacc (���� - SF ����� �������, �������� - ������ � ������� SFacc)

    Dim accSF(0 To 9999) As String, accSFind(0 To 9999) As String
    hashInit accSF, accSFind
    For i = 2 To EOL_SFacc
        hashSet accSF, accSFind, Compressor(Sheets(SFacc).Cells(i, SFACC_ACC1C_COL)), CInt(i)
    Next i

'   2. ������� SFacc (���� - ����� �� ����� ����������� SF,
'                     �������� - ������ � SFacc, ������������� �������� '$'

    Dim accSFComps(0 To 9999) As String, accSFIndxs(0 To 9999) As String
    
    hashInit accSFComps, accSFIndxs
    For i = 2 To EOL_SFacc
        x = split(LCase$(RemIgnored(Trim$(Sheets(SFacc).Cells(i, SFACC_ACCNAME_COL)))))
        For j = 0 To UBound(x)
            s0 = hashGet(accSFComps, accSFIndxs, x(j))      ' ��������� ���������
            If s0 <> "$" Then                               ' �������� ��� ������ account'a
                s0 = s0 + "$"                               ' ����������� "$"
            Else
                s0 = ""
            End If
            hashSet accSFComps, accSFIndxs, x(j), s0 + Str(i)
        Next j
    Next i

'---------- ������ �� ����������� �������� 1� --------------------------

' EOL_AccntUpd - ������ ��������� ����� - ����������� ����������� SF
' EOL_AdAcc    - ������ ��������� ����� - �������� ����������� SF
    
'    ProgressForm.Show vbModeless
'    ProgressForm.ProgressLabel.Caption = Doing

' ���� ���� SF �� ������ 1�

    Dim accntName As String, accntNamePrev As String, wrSF As String
    Dim accntAddr As String
    Dim sfWrds() As String, SFWordIndx As Long
    Dim Msg As String
    Dim CompSNums(1 To 100) As Long, compNum As Long                ' �����a ��������� �����
    Dim namSF(1 To 100) As String, count(1 To 100) As Long          ' SFacc ������ ���
    Dim adrTxt(1 To 100) As String, kword(1 To 100) As String
    Dim adrField As String
    Dim AdrStruct As PostAddr, AdrStruct1C As PostAddr
    Dim delAddrSF As PostAddr, factAddr1C As PostAddr
    Dim DlgRes As String

    For i = 2 To EOL_Acc1C                                          ' ���� �� ����������� 1�
'    For i = 1507 To EOL_Acc1C                                          ' ���� �� ����������� 1�
        If ExRespond = False Then GoTo ExitSub
        
        With Sheets(Acc1C)
            accntAddr = .Cells(i, A1C_ADR_COL)
                    
            If accntAddr <> "" Then ' ���� ���� '�����' (�������� �����) �� ���������, ����������
            
                accntName = .Cells(i, A1C_NAME_COL)
                ' ������������ �� 1� �����. �������, ��� ������������� �� ����� �����������
                If accntName <> accntNamePrev Then
                   
                    ' ���������, ��� �� ��� ������ �� ������-������ SF account'a �� ��� ����������� 1�
                    If hashGet(accSF, accSFind, Compressor(accntName)) = "$" Then
                        
                        Msg = Str(i) + ":  ��� 1�:     " + accntName + vbCrLf _
                            + "�����:              " + .Cells(i, A1C_ADR_COL)
'                        MSG2 = ""
                        compNum = 0
                        ' ��������� 1� ��� �� ����� � ��������� ������: ���� ����� � ���� ���� SF
                        sfWrds = split(LCase$(RemIgnored(accntName)))
                        For SFWordIndx = 0 To UBound(sfWrds)
                            wrSF = hashGet(accSFComps, accSFIndxs, sfWrds(SFWordIndx))
                            If wrSF <> "$" Then
                                x = split(wrSF, "$")
                                For j = 0 To UBound(x)
                                    adrField = SFPostAddr(x(j), SFacc)
                                    If adrField <> "" Then    ' ���������� ������ ��� ������
                                        compNum = compNum + 1           ' ������� ��������
                                        CompSNums(compNum) = x(j)       ' ���������� ����� ������ SFacc
                                        namSF(compNum) = Sheets(SFacc).Cells(x(j), SFACC_ACCNAME_COL)
                                        adrTxt(compNum) = adrField
                                        
                                        kword(compNum) = sfWrds(SFWordIndx)
                                   End If
                                Next j
                            End If
                        Next SFWordIndx
                        
        ' namSF - ������ ���� � SF, adrTxt - ������ ��������������� �������,
        ' CompSNums - ��������������� ������ ����� � SFacc
        ' compNum - ����� ���������
        
        ' ������������ � ���������� �� ���������� ����������.
        ' ��� ������ ����������, ��� ���� ��������� � ��� ����� � ������.
        
                        For j = 1 To compNum        ' ������������ ��������� ����������
                            count(j) = 1
                        Next j
                        For j = 1 To compNum - 1       ' ����� ����������
                            For k = j + 1 To compNum
                                If k <> j Then
                                    If namSF(k) = namSF(j) Then
                                        count(j) = count(j) + 1     ' ����� ��������
                                        count(k) = -1000            ' �������� �������� �� ��������
                                        kword(j) = kword(j) + " " + kword(k)
                                    End If
                                End If
                            Next k
                        Next j
                        
'                   ���������� �� ����� ���������� (���������).
'                   ���� ��������� � count = -1000 ���������� � ����� � �� ���������� � �����
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
        
'                   ��������� ���������� ������
                        For j = 1 To compNum
                            If count(j) > 0 Then
'                                MSG2 = MSG2 + vbCrLf + vbCrLf + "            " _
'                                    + Format(j) + ".     '" + "'" + namSF(j) + "'            �����: " + kword(j) _
'                                    + vbCrLf & "                 �����: " + adrTxt(j)
                            Else
                                compNum = j - 1             ' ������������ ����������� ��� �������
                                GoTo endLoopPrepTxt         ' � ��������� ����� �� �����
                            End If
                        Next j
endLoopPrepTxt:
        
'                   ����� �����������. ��������� ������.
                        Dim Repeat As Boolean
                        Do
                            Repeat = False
                            DlgRes = DlgAccChoice(CompSNums, compNum, A1C_NAME_COL, Msg, namSF, adrTxt, kword)
                            If IsNumeric(DlgRes) Then  ' SF account id  + 1C id
    '                            MsgBox "������� " + DlgRes + vbCrLf + " (" _
    '                                + Sheets(Acc1C).Cells(i, A1C_NAME_COL) + "; " _
    '                                + Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_ACCNAME_COL) + "')" _

    
                                
    
    
                            ' ��������� � �������� �����
                                SFaccMergeWith1C.SFacc = accntName
                                SFaccMergeWith1C.name1C = namSF(CInt(DlgRes))
                                SFaccMergeWith1C.setInn Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_INN_COL), _
                                                        Sheets(Acc1C).Cells(i, A1C_INN_COL)
                                SFaccMergeWith1C.setTel Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_TEL_COL), _
                                                        Sheets(Acc1C).Cells(i, A1C_TEL_COL)
                                                        
                                ' ���������� �������� ����� �����
                                AdrStruct.City = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_CITY_COL)
                                AdrStruct.Street = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_STREET_COL)
                                AdrStruct.State = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_STATE_COL)
                                AdrStruct.PostIndex = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_INDEX_COL)
                                AdrStruct.Country = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_COUNTRY_COL)
                                AdrStruct1C = AdrParse(Sheets(Acc1C).Cells(i, A1C_ADR_COL))
                                delAddrSF.City = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_DELCITY_COL)
                                delAddrSF.Street = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_DELSTREET_COL)
                                delAddrSF.State = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_DELSTATE_COL)
                                delAddrSF.PostIndex = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_DELINDEX_COL)
                                delAddrSF.Country = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_DELCOUNTRY_COL)
                                factAddr1C = AdrParse(Sheets(Acc1C).Cells(i, A1C_FACTADR_COL))
                                SFaccMergeWith1C.setAddr AdrStruct, AdrStruct1C, delAddrSF, factAddr1C
                               
                                SFaccMergeWith1C.setTel Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_TEL_COL), _
                                                        Sheets(Acc1C).Cells(i, A1C_TEL_COL)
                                
                                SFaccMergeWith1C.Show                               ' ����� �����
                                
                                If SFaccMergeWith1C.result = "exit" Then            ' ��������� ����������� �����
                                    ExRespond = False
                                ElseIf SFaccMergeWith1C.result = "save" Then
                                    EOL_AccntUpd = EOL_AccntUpd + 1
                                    With Sheets(AccntUpd)
                                        .Cells(EOL_AccntUpd, ACCUPD_SFID_COL) = Sheets(SFacc).Cells(CompSNums(CInt(DlgRes)), SFACC_IDACC_COL)
                                        .Cells(EOL_AccntUpd, ACCUPD_1CNAME_COL) = accntName     ' ��� �� ����������� 1� -> SF account
                                        .Cells(EOL_AccntUpd, ACCUPD_INN_COL) = SFaccMergeWith1C.innSF
                                        .Cells(EOL_AccntUpd, ACCUPD_TEL_COL) = SFaccMergeWith1C.telSF
                                    End With
                                ElseIf SFaccMergeWith1C.result = "back" Then
                                    Repeat = True       ' ������������ ������ ���������� ����������
                                End If  '���� �� ���� ������� �� ��������� - ������ "����������'
                            ElseIf DlgRes = "create" Then
                                ' ��������� ���� �����: ��� 1�, ��� SF (������ ��� ���������)
                                NewSFaccForm.Adr1C.Caption = Sheets(Acc1C).Cells(i, A1C_NAME_COL)
                                NewSFaccForm.SFacc.value = NewSFaccForm.Adr1C.Caption
                                
                                ' �������� �����
                                s0 = Trim(Sheets(Acc1C).Cells(i, A1C_ADR_COL))
                                AdrStruct = AdrParse(s0)
                                NewSFaccForm.Area.value = AdrStruct.State
                                NewSFaccForm.City.value = AdrStruct.City
                                NewSFaccForm.Street.value = AdrStruct.Street
                                NewSFaccForm.Index.value = AdrStruct.PostIndex
                                NewSFaccForm.Country.value = AdrStruct.Country
                                ' ����������� �����(1�) / ����� ��������(SF)
                                s0 = Trim(Sheets(Acc1C).Cells(i, A1C_FACTADR_COL))
                                AdrStruct = AdrParse(s0)
                                NewSFaccForm.AreaD.value = AdrStruct.State
                                NewSFaccForm.CityD.value = AdrStruct.City
                                NewSFaccForm.StreetD.value = AdrStruct.Street
                                NewSFaccForm.IndexD.value = AdrStruct.PostIndex
                                NewSFaccForm.CountryD.value = AdrStruct.Country
                                
                                NewSFaccForm.contact.value = Sheets(Acc1C).Cells(i, A1C_CON_COL)
                                Dim INN
                                INN = Trim(Sheets(Acc1C).Cells(i, A1C_INN_COL))
                                If INN <> "" Then INN = split(INN, "/")(0)
                                NewSFaccForm.INN = Trim(INN)
                                NewSFaccForm.phone.value = Sheets(Acc1C).Cells(i, A1C_TEL_COL)
                                
                                ' ���������� ���� - �� ��������
                                NewSFaccForm.invoice.Caption = Sheets(Acc1C).Cells(i, A1C_INVOICE_COL)
                                NewSFaccForm.good.Caption = Sheets(Acc1C).Cells(i, A1C_GOOD_COL)
                                
                                NewSFaccForm.Show vbModal
                                
                                DlgRes = NewSFaccForm.result.value
                                If DlgRes = "exit" Then
                                    ExRespond = False
                                ElseIf DlgRes = "save" Then
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
                                
                            End If      ' ����� isnumeric()
                        Loop While Repeat
                    End If              ' ����� if �� ������� � SF
                End If                  ' ����� if �� ������������
            End If                      ' ����� if �� ���� '1� �����
        End With
    Next i
ExitSub:
    
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Account\"
    WriteCSV A_Acc, "AdAcc.txt"
    Shell "quota_Acc.bat"
    
    WriteCSV AccntUpd, "AccntUpd.txt"
    Shell "quotaAccUpd.bat"
    
    MS "����: created " + Str(EOL_AdAcc) + " linked " + Str(EOL_AccntUpd)
End Sub

Function SFPostAddr(ByVal indx As Long, SFacc As String)
'   ����������� ������������� ��������� ������
' 31.05.12

    With Sheets(SFacc)
        SFPostAddr = Replace((.Cells(indx, SFACC_CITY_COL) _
                + "," + .Cells(indx, SFACC_STATE_COL) _
                + "," + .Cells(indx, SFACC_STREET_COL) _
                + "," + .Cells(indx, SFACC_INDEX_COL) _
                + "," + .Cells(indx, SFACC_COUNTRY_COL)), ",,", ",")
    End With
End Function
Function DlgAccChoice(CompSNums, count, idCol, Msg, namSF, addrTxt, kword)
    ' CompSNums - ������ ������� ����� � �������
    ' count     - actual possibility count
    ' idCol     - ����� ������� � �������
    ' MSG       - ��������� ������� � ���������, �����, �� ��������� �� ������
    ' namSF     - ����� �����������
    ' addrTxt   - ������ �����������
    ' kword     - �������� ����� �� ������� ����������� �������
    
    Dim i As Long
    
    If count = 0 Then
        DlgAccChoice = "create"       ' the only possibility
        Exit Function
    End If
    
    DlgAccChoice = "cont"       ' ���� ��������� �� ����� - ��������� ���
    SFaccountForm.accntChoice.ColumnCount = 3
    Do                          ' ���� �� �����������, ������� ����� �������
        ' ������� listbox ������
        Do While SFaccountForm.accntChoice.ListCount <> 0
            SFaccountForm.accntChoice.RemoveItem 0
        Loop
        ' ���������� listbox
        For i = 1 To count
            SFaccountForm.accntChoice.AddItem
            SFaccountForm.accntChoice.List(i - 1, 0) = namSF(i)
            SFaccountForm.accntChoice.List(i - 1, 1) = addrTxt(i)
            SFaccountForm.accntChoice.List(i - 1, 2) = kword(i)
        Next i
                    
        SFaccountForm.TextBox2 = Msg                        ' �������� �����: ��� SF, ����� SF
        If count = 1 Then
            SFaccountForm.TextBox1.value = "1"              ' ���� ������ ���, ����������� default
            SFaccountForm.accntChoice.ListIndex = 0         ' listbox - ������� ������������ ������
        Else
            SFaccountForm.accntChoice.ListIndex = -1        ' listbox - �� �������
            SFaccountForm.TextBox1.value = ""               ' �������� �������� ������ - �����
        End If
        
        ' textbox �������, �� ("�������") - ���� ���� ����������� �������
        SFaccountForm.OKButton.Visible = True
        If count = 0 Then SFaccountForm.OKButton.Visible = False
        
        SFaccountForm.Show vbModal                      ' ������ � ������
        
        Dim inpt As String, j As Long
        inpt = SFaccountForm.TextBox1
        j = 0                                           ' �� ������ ������������� �����, � ����. account'�
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
        If MsgBox("���������� ������� �����������. ����������?", vbYesNo) <> vbYes Then Exit Do
    Loop
endloop:
    If inpt = "exit" Then ExRespond = False

End Function
Function switch(kword, j, k)

' 2 ���������� ������� �������� �������
' 5.6.2012
    Dim s As String
    s = kword(j)
    kword(j) = kword(k)
    kword(k) = s
End Function


