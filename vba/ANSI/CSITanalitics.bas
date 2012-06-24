Attribute VB_Name = "CSITanalitics"
'----------------------------------------------------------
' ������ ������� ���������� �� CSIT
'   ���� �.     20.6.12
'(?)CSIT_MS_Clear()                             - ������� ��������� ����� MS_CSIT
'(?)CSIT_MS_lead()                              - ������ ���� ����������� - ����� Microsoft CSIT
'(*)PaymntCl1CAnlz()                            - ��������� �������� 1� �� ����� "���� ����� ��������"
'(*)Client1CAnlz                                - ��������� �������� 1� �� ����� "������ �������� 1C"
' - client1CProcess(ByVal accntName As String)  - ��������� ������ ������� 1�
' - SFPostAddr(indx As Long, SFacc As String)   - ����������� ������������� ��������� ������
' - DlgAccChoice                                - ����� ����� "����� ����������� SF ��� ����������"
' T testTelToFax()                              - ���� ��� telToFax
' - telToFax(tel)                               - ����� ������� ������ � ������ ������� ���������

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
                        SFaccountForm.TextBox2 = Msg + MSG2       ' �������� �����
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
Sub PaymntCl1CAnlz()

'   ��������� �������� 1� �� ����� "��������"
'       22.06.12

    Dim i As Long, j As Long, k As Long

    ModStart PAY_SHEET, "������ ���� ����� ��������"
    
    CheckSheet SFacc, EOL_SFacc + 2, 1, SFaccRepName
    CheckSheet PAY_SHEET, 1, PAYDOC_COL, Stamp1Cpay1
    
    ClearSheet A_Acc, Range("HDR_AdAcc")
    ClearSheet AccntUpd, Range("HDR_AccntUpd")
    
    For i = 2 To EOL_PaySheet
        If ExRespond = False Then GoTo ExitSub
'If InStr(LCase$(Sheets(PAY_SHEET).Cells(i, PAYACC_COL)), "����") <> 0 Then
'    i = i
'End If
        client1CProcess Sheets(PAY_SHEET).Cells(i, PAYACC_COL)      ' �������� - ��� ������� 1�
    Next i
ExitSub:
    MS "����: ������� " + (Str(EOL_AdAcc - 1)) + "; ������� " + (Str(EOL_AccntUpd - 1))
    
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Account\"
    WriteCSV A_Acc, "AdAcc.txt"
    Shell "quota_Acc.bat"
    
    WriteCSV AccntUpd, "AccntUpd.txt"
    Shell "quotaAccUpd.bat"
End Sub
Sub Client1CAnlz()

'   ��������� �������� 1� �� ����� "������ �������� 1C"
'       22.06.12

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
    MS "����: ������� " + (Str(EOL_AdAcc - 1)) + "; ������� " + (Str(EOL_AccntUpd - 1))
    
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Account\"
    WriteCSV A_Acc, "AdAcc.txt"
    Shell "quota_Acc.bat"
    
    WriteCSV AccntUpd, "AccntUpd.txt"
    Shell "quotaAccUpd.bat"
End Sub

Sub client1CProcess(ByVal accntName As String)

'   ��������� ������ ������� 1�
'   accntName - ��� �������
'       19.06.12

    Dim s0 As String, S1() As String                    ' ��������� ����������
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
'If InStr(LCase$(Sheets(SFacc).Cells(i, SFACC_ACC1C_COL)), "����") <> 0 Then
'    i = i
'End If
            hashSet accSF, accSFind, Compressor(Sheets(SFacc).Cells(i, SFACC_ACC1C_COL)), CInt(i)
        Next i
    
    '   2. ������� SFacc (���� - ����� �� ����� ����������� SF,
    '                     �������� - ������ � SFacc, ������������� �������� '$'
    
        
        hashInit accSFComps, accSFIndxs
        For i = 2 To EOL_SFacc
            S1 = split(LCase$(RemIgnored(Trim$(Sheets(SFacc).Cells(i, SFACC_ACCNAME_COL)))))
            For j = 0 To UBound(S1)
                s0 = hashGet(accSFComps, accSFIndxs, S1(j))     ' ��������� ���������
                If s0 <> "$" Then                               ' �������� ��� ������ account'a
                    s0 = s0 + "$"                               ' ����������� "$"
                Else
                    s0 = ""
                End If
                hashSet accSFComps, accSFIndxs, S1(j), s0 + Str(i)
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
    
' --------- ��������� ������� 1� -----------------
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

    s0 = hashGet(acc1CHashKey, acc1CHashVal, LCase$(accntName))     ' ��� ���� ������ � ������� 1�
    If s0 = "$" Then
                                                                    ' false - ��� stop'� � ������� "����������?"
        ErrMsg TYPE_ERR, "client1CProcess: ������� '" + accntName + "' ��� � ����������� �������� 1�", False
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
            Msg = "��� 1�:     " + accntName _
                + vbCrLf + "�����:              " + .Cells(clIndx, A1C_ADR_COL)
'                        MSG2 = ""
            compNum = 0
            ' ��������� 1� ��� �� ����� � ��������� ������: ���� ����� � ���� ���� SF
            sfWrds = split(LCase$(RemIgnored(accntName)))
            For SFWordIndx = 0 To UBound(sfWrds)
                wrSF = hashGet(accSFComps, accSFIndxs, sfWrds(SFWordIndx))
                If wrSF <> "$" Then
                    S1 = split(wrSF, "$")
                    For j = 0 To UBound(S1)
                        adrField = SFPostAddr(S1(j), SFacc)
                        If adrField <> "" Then                  ' ���������� ������ ��� ������
                            compNum = compNum + 1               ' ������� ��������
                            CompSNums(compNum) = S1(j)          ' ���������� ����� ������ SFacc
                            namSF(compNum) = Sheets(SFacc).Cells(S1(j), SFACC_ACCNAME_COL)
                            adrTxt(compNum) = adrField
                            
                            kword(compNum) = sfWrds(SFWordIndx) '�������� �����
                       End If
                    Next j
                End If
            Next SFWordIndx
        End With
                        
        ' namSF - ������ ���� � SF, adrTxt - ������ ��������������� �������,
        ' CompSNums - ��������������� ������ ����� � SFacc
        ' compNum - ����� ���������
        
        ' ������������ �� SF account'�� � ���������� �� ���������� ����������.
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
        Dim Repeat As Boolean               ' ���� ��������� ��������� account'a 1C
        Dim setLink As Boolean              ' ���� ���������� account'a 1C � ���������� account'�� SF.
                                            ' ����� ���� ������� ��� ������������ ������ ���������
                                            ' �� �������������� "��� ����������� ��� ������� � SF!"
        Do
            Repeat = False                  ' ���� �� ���������, ������ �� ����� Do
            DlgRes = DlgAccChoice(CompSNums, compNum, A1C_NAME_COL, Msg, namSF, adrTxt, kword)
            If IsNumeric(DlgRes) Then  ' SF account id  + 1C id
            
                setLink = True
                If Sheets(SFacc).Cells(CompSNums(DlgRes), SFACC_ACC1C_COL) <> "" Then
                            ' ��� ����������� ��� �������!
                    If MsgBox("��� ����������� ��� ������� � SF!" _
                        + vbCrLf + "�� ������������� ������ �������� �����?", vbYesNo) <> vbYes Then
                        setLink = False
                        Repeat = True
                    End If
                End If
                If setLink Then
                    ' ��������� � �������� �����
                    SFaccMergeWith1C.SFacc = accntName
                    SFaccMergeWith1C.name1C = namSF(CInt(DlgRes))
                    j = CompSNums(CInt(DlgRes))
                    SFaccMergeWith1C.setInn Sheets(SFacc).Cells(j, SFACC_INN_COL), _
                                            Sheets(Acc1C).Cells(clIndx, A1C_INN_COL)
                    SFaccMergeWith1C.setTel Sheets(SFacc).Cells(j, SFACC_TEL_COL), _
                                            Sheets(Acc1C).Cells(clIndx, A1C_TEL_COL)
            
                    ' ���������� �������� ����� �����
                    With Sheets(SFacc)
                        AdrStruct.City = .Cells(j, SFACC_CITY_COL)
                        AdrStruct.Street = .Cells(j, SFACC_STREET_COL)
                        AdrStruct.State = .Cells(j, SFACC_STATE_COL)
                        AdrStruct.PostIndex = .Cells(j, SFACC_INDEX_COL)
                        AdrStruct.Country = .Cells(j, SFACC_COUNTRY_COL)
                        delAddrSF.City = .Cells(j, SFACC_DELCITY_COL)
                        delAddrSF.Street = .Cells(j, SFACC_DELSTREET_COL)
                        delAddrSF.State = .Cells(j, SFACC_DELSTATE_COL)
                        delAddrSF.PostIndex = .Cells(j, SFACC_DELINDEX_COL)
                        delAddrSF.Country = .Cells(j, SFACC_DELCOUNTRY_COL)
                    End With
                    With Sheets(Acc1C)
                        AdrStruct1C = AdrParse(.Cells(clIndx, A1C_ADR_COL))
                        factAddr1C = AdrParse(.Cells(clIndx, A1C_FACTADR_COL))
                    End With
                    SFaccMergeWith1C.setAddr AdrStruct, AdrStruct1C, delAddrSF, factAddr1C
            
                    SFaccMergeWith1C.setTel Sheets(SFacc).Cells(j, SFACC_TEL_COL), _
                                            Sheets(Acc1C).Cells(clIndx, A1C_TEL_COL)
            
                    SFaccMergeWith1C.Show                               ' ����� �����
            
                    If SFaccMergeWith1C.result = "exit" Then            ' ��������� ����������� �����
                        ExRespond = False
                    ElseIf SFaccMergeWith1C.result = "save" Then
                    
                        EOL_AccntUpd = EOL_AccntUpd + 1
                        With Sheets(AccntUpd)
                            .Cells(EOL_AccntUpd, ACCUPD_SFID_COL) = Sheets(SFacc).Cells(j, SFACC_IDACC_COL)
                            .Cells(EOL_AccntUpd, ACCUPD_1CNAME_COL) = accntName     ' ��� �� ����������� 1� -> SF account
                            .Cells(EOL_AccntUpd, ACCUPD_INN_COL) = SFaccMergeWith1C.innSF
                            .Cells(EOL_AccntUpd, ACCUPD_TEL_COL) = SFaccMergeWith1C.telSF
                            .Cells(EOL_AccntUpd, ACCUPD_FAX_COL) = SFaccMergeWith1C.faxSF
                            .Cells(EOL_AccntUpd, ACCUPD_CITY_COL) = SFaccMergeWith1C.CitySF
                            .Cells(EOL_AccntUpd, ACCUPD_STREET_COL) = SFaccMergeWith1C.StreetSF
                            .Cells(EOL_AccntUpd, ACCUPD_STATE_COL) = SFaccMergeWith1C.AreaSF
                            .Cells(EOL_AccntUpd, ACCUPD_INDEX_COL) = SFaccMergeWith1C.IndexSF
                            .Cells(EOL_AccntUpd, ACCUPD_STATE_COL) = SFaccMergeWith1C.CountrySF
                            .Cells(EOL_AccntUpd, ACCUPD_COUNTRY_COL) = SFaccMergeWith1C.AreaSF
                            .Cells(EOL_AccntUpd, ACCUPD_DELCITY_COL) = SFaccMergeWith1C.DelCitySF
                            .Cells(EOL_AccntUpd, ACCUPD_DELSTREET_COL) = SFaccMergeWith1C.DelStreetSF
                            .Cells(EOL_AccntUpd, ACCUPD_DELSTATE_COL) = SFaccMergeWith1C.DelAreaSF
                            .Cells(EOL_AccntUpd, ACCUPD_DELINDEX_COL) = SFaccMergeWith1C.DelIndexSF
                            .Cells(EOL_AccntUpd, ACCUPD_DELCOUNTRY_COL) = SFaccMergeWith1C.DelCountrySF
                        End With
                    ElseIf SFaccMergeWith1C.result = "back" Then
                        Repeat = True       ' ������������ ������ ���������� ���������� ����� do
                    End If  '���� �� ���� ������� �� ��������� - ������ "����������'
                End If      '����� if setLink
            ElseIf DlgRes = "create" Then
                ' ��������� ���� �����: ��� 1�, ��� SF (������ ��� ���������)
                NewSFaccForm.Adr1C.Caption = Sheets(Acc1C).Cells(clIndx, A1C_NAME_COL)
                NewSFaccForm.SFacc.value = NewSFaccForm.Adr1C.Caption
        
                ' �������� �����
                s0 = Trim(Sheets(Acc1C).Cells(clIndx, A1C_ADR_COL))
                NewSFaccForm.setPostAddr AdrParse(s0)
                ' ����������� �����(1�) / ����� ��������(SF)
                s0 = Trim(Sheets(Acc1C).Cells(clIndx, A1C_FACTADR_COL))
                NewSFaccForm.setDelAddr AdrParse(s0)
        
                NewSFaccForm.contact.value = Sheets(Acc1C).Cells(clIndx, A1C_CON_COL)
                Dim INN
                INN = Trim(Sheets(Acc1C).Cells(clIndx, A1C_INN_COL))
                If INN <> "" Then INN = split(INN, "/")(0)
                NewSFaccForm.INN = Trim(INN)
                NewSFaccForm.phone.value = Sheets(Acc1C).Cells(clIndx, A1C_TEL_COL)
        
            ' ���������� ���� - �� �������� ����������
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
                        .Cells(EOL_AdAcc, ADACC_CITY_COL) = NewSFaccForm.City
                        .Cells(EOL_AdAcc, ADACC_STATE_COL) = NewSFaccForm.Area
                        .Cells(EOL_AdAcc, ADACC_STREET_COL) = NewSFaccForm.Street
                        .Cells(EOL_AdAcc, ADACC_INDEX_COL) = NewSFaccForm.Index
                        .Cells(EOL_AdAcc, ADACC_COUNTRY_COL) = NewSFaccForm.Country
                        .Cells(EOL_AdAcc, ADACC_CONTACT1C_COL) = NewSFaccForm.contact
                        .Cells(EOL_AdAcc, ADACC_INN_COL) = NewSFaccForm.INN
                        .Cells(EOL_AdAcc, ADACC_TEL_COL) = NewSFaccForm.phone
                        .Cells(EOL_AdAcc, ADACC_FAX_COL) = NewSFaccForm.fax
                        .Cells(EOL_AdAcc, ADACC_FACTCITY_COL) = NewSFaccForm.CityD
                        .Cells(EOL_AdAcc, ADACC_FACTSTATE_COL) = NewSFaccForm.AreaD
                        .Cells(EOL_AdAcc, ADACC_FACTSTREET_COL) = NewSFaccForm.StreetD
                        .Cells(EOL_AdAcc, ADACC_FACTINDEX_COL) = NewSFaccForm.IndexD
                        .Cells(EOL_AdAcc, ADACC_FACTCOUNTRY_COL) = NewSFaccForm.CountryD
                    End With
                ElseIf DlgRes = "back" Then
                    Repeat = True
                End If      ' Dlgres= 'exit'
            End If          ' isnumeric(dlgres)
        Loop While Repeat
    End If                  'end if �� ������������
    
exitProc:
End Sub

Function SFPostAddr(ByVal indx As Long, SFacc As String)
'   ����������� ������������� ��������� ������ - ���� ����� ������� (�����, �������, �����/���, ������, ������)
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
    
    SFaccountForm.TextBox2 = Msg        ' ��� � ����� ����������� 1�
    
    
    NewSFaccForm.BackButton.Enabled = True
    If count = 0 Then
        NewSFaccForm.BackButton.Enabled = False ' ����� ��������� � ������ ����� ����������, ������ ���������
        DlgAccChoice = "create"         ' the only possibility
        Exit Function
    End If
    
    DlgAccChoice = "cont"       ' ���� ��������� �� ����� - ��������� ���
    SFaccountForm.accntChoice.ColumnCount = 4
    Do                          ' ���� �� �����������, ������� ����� �������
        ' ������� listbox'� �������
        Do While SFaccountForm.accntChoice.ListCount <> 0
            SFaccountForm.accntChoice.RemoveItem 0
        Loop
        
        ' ���������� listbox'��
        For i = 1 To count
            SFaccountForm.accntChoice.AddItem
            SFaccountForm.accntChoice.List(i - 1, 0) = ""
            If Sheets(SFacc).Cells(CompSNums(i), SFACC_ACC1C_COL) <> "" Then
                SFaccountForm.accntChoice.List(i - 1, 0) = "�������"     ' ��� ����������� ��� �������!
            End If
            SFaccountForm.accntChoice.List(i - 1, 1) = namSF(i)
            SFaccountForm.accntChoice.List(i - 1, 2) = addrTxt(i)
            SFaccountForm.accntChoice.List(i - 1, 3) = kword(i)
        Next i
             ' �������� �����: ��� SF, ����� SF
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
Sub testTelToFax()
' ���� ��� telToFax
    Dim t1, t2, t3, t4, t5, t6, t7, t8, t9, t10
    t1 = telToFax("1234 f c 1234-45(55)")
    t2 = telToFax("1234 f  1234-45(55)")
    t3 = telToFax("1234 fax1234-45(55)")
    t4 = telToFax("1234 fax 1234-45(55)")
    t5 = telToFax("1234 � 1234-45(55)???????????????????  ���� 7(33)444")
    t6 = telToFax("1234 fax +1234-45(55)")
    t7 = telToFax("1234 fax 1234-45(556)")
    
    If t1 <> "" Or t2 <> "���� 1234-45(55)" Or t3 <> "���� 1234-45(55)" _
            Or t4 <> "���� 1234-45(55)" Or t5 <> "���� 1234-45(55),���� 7(33)444" _
            Or t6 <> "���� 1234-45(55)" Or t7 <> "���� 1234-45(556)" Then
        Stop ' ������!
    End If
    ' ���� �� ����� - ���� ������
End Sub
Function telToFax(tel)
' ����� ������� ������ � ������ ������� ���������
'   22.06.12

    Dim i As Long, j As Long
    Dim sym As String, rest As String
    Dim beg As Long
    Dim pref() As String
    pref = split("fax.;fax;f.;f;����.;����;�.;�", ";")  ' �������� ����� (������� �������)
    
    
    telToFax = ""
        
    For i = 0 To Len(tel)
        beg = -1                            ' ���� �� ������ �������� ����� - ��������� -1
        For j = LBound(pref) To UBound(pref)
            sym = LCase(Mid(tel, i + 1, Len(pref(j))))  ' � Mid ��������� �������� � 1, � �� � 0, ������� i+1
            If sym = pref(j) Then
                beg = 0
                i = i + Len(pref(j))            ' ���������� ���������
                GoTo jBreak
            End If
        Next j
jBreak:
        If beg = 0 Then
            For j = i To Len(tel)               ' ����� �������� �����, ���� ����� �����
                sym = LCase(Mid(tel, j + 1, 1)) ' ������� ������
                If sym <> " " And sym <> "+" Then   ' ���������� " " & "+"
                                                ' � ������ ������ ��������� ����� ��� ������
                    If IsNumeric(sym) Or sym = "(" Or sym = ")" _
                            Or (beg <> 0 And sym = "-") Then    ' � �������� - ��� '-'
                        If beg = 0 Then beg = j
                    ElseIf beg = 0 Then
                        GoTo endSub             ' ������������ ������ � ������
                    Else                        ' ������������ ������ - ����������
                        telToFax = "���� " + Mid(tel, beg + 1, j - beg) ' ��������� ���������
                        rest = telToFax(Mid(tel, j + 1, 999))           ' �������� - ���� ������ ������ � ������� ������
                        If rest <> "" Then
                            telToFax = telToFax + "," + rest            ' ���� - ��������� � ��������� ����� �������
                        End If
                        GoTo endSub                                     ' ��� �������
                    End If
                End If
            Next j
        End If
    Next i
endSub:
End Function

Function switch(kword, j, k)

' 2 ���������� ������� �������� �������
' 5.6.2012
    Dim S As String
    S = kword(j)
    kword(j) = kword(k)
    kword(k) = S
End Function


