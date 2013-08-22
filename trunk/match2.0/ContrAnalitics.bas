Attribute VB_Name = "ContrAnalitics"
'-----------------------------------------------------------------------------
' ContrAnalitics    - ������ ���������� � ����������� � ��� ��������
' [*] ContrPass     - ������ �� ����� ��������� (3) � �������� ����� ��������
'  -  OppT(V1C)     - ���������� ��� ������� �� ���� ��������� � �������� 1�
' [*] DogOppLink    - ������ �� SFD � ����� ���������� �������� ��� �����
'  -  IsSameVendor(OppType, V1C, ContrCode)    - ���������� True, ���� ����&���
'                           ������� OppType ������������� ���������� �� �������� � 1�
'   3.2.2013

Option Explicit
Sub NewContr(NewContract As String)
'
' S NewContr()  - �������� ��������� 1� ��� ��������� � SF ����� ����� DL
' 18.8.2012
'  3.9.12 - StepIn
' 31.10.10 - fix Call WrNewSheet
'  9.11.12 - ����� NewSheet �� NewContr

    StepIn
    
    Dim Dog As TOCmatch
    Dim i As Long
    
    NewSheet NewContract
    
    Dog = GetRep(DOG_SHEET)

    With DB_1C.Sheets(DOG_SHEET)
        For i = 2 To Dog.EOL
            Progress i / Dog.EOL
            If .Cells(i, DOGIDSF_COL) = "" And .Cells(i, DOGISACC_COL) <> "" Then
                WrNewSheet NewContract, DOG_SHEET, i
            End If
        Next i
    End With
End Sub
Sub PaidContr(ByVal NewPayment As String)
'
' S PaidContr()  - ��������� ������� � ���������, ��������� � ��������
' 21.11.12
' 30.12.12 - NewOpp ��� ��������� ��� ������ � ��������

    StepIn
    
    Dim Paid As TOCmatch, ContrK As String, IsErr As Boolean, OppId As String
    Dim i As Long
    
    Const F_rqst = "SFD/" & SFD_COD_COL & ":" & SFD_OPPID_COL
    
    NewSheet NewPayment
    
    Paid = GetRep(PAY_SHEET)

    With DB_1C.Sheets(PAY_SHEET)
        For i = 2 To Paid.EOL
            Progress i / Paid.EOL
            If .Cells(i, PAYINSF_COL) <> 1 And .Cells(i, PAYDOGOVOR_COL) <> "" Then
                ContrK = ContrCod(.Cells(i, PAYDOGOVOR_COL), .Cells(i, PAYOSNDOGOVOR_COL))
                OppId = FetchDoc(F_rqst, ContrK, IsErr)
                If OppId = "" Then
'---- NEWOPP              ���� ��� CSV    ����      ������ -- �� �������� ���!!!
'                    WrNewSheet NEW_OPP, PAY_SHEET, i, "HDR_NewOppBy"
                End If
                If Not IsErr Then WrNewSheet NewPayment, PAY_SHEET, i, OppId
            End If
        Next i
    End With
End Sub
Sub ContrPass()
'
' ������ �� ������ ��������� � �����/�������� ��������������� ��������
'   4.4.12
    
    Dim Contr As Range
    Dim i, OppN, OppIs, OppNew, NoOpp As Integer
    Dim IsGenOpp As Integer
    Dim DogSFstat, ContrK, ContrId, OppTyp, t, Stage As String
    Dim Acc, Dat, Sale, DogValue, DogCur, OppStage As String
    Dim VendorSF, Vendor1C As String
    
    Lines = ModStart(DOG_SHEET, "������ �� ���������: ��������� �������", True) - DOGRES
    EOL_DogSheet = Lines
    EOL_SFD = EOL(SFD) - SFresLines
    EOL_SFopp = EOL(SFopp) - SFresLines
    EOL_SFacc = EOL(SFacc) - SFresLines
    
    CheckSheet DOG_SHEET, 1, 10, Stamp1Cdog1
    CheckSheet SFD, EOL_SFD + 2, 3, SFcontrRepName
    CheckSheet SFopp, EOL_SFopp + 2, 1, SFoppRepName
'---------- ������ �� ��������� ------------------------
    OppIs = 0: OppNew = 0: NoOpp = 0: Fruitful = 0
    ClearSheet O_NewOpp, Range("HDR_NewOpp")
    ClearSheet NewContract, Range("HDR_NewContract")
    ClearSheet NewContractLnk, Range("HDR_ContrLnk")
    
    For i = 2 To Lines
        Progress i / Lines
        If ExRespond = False Then Exit For
        With Sheets(DOG_SHEET)
            DogSFstat = .Cells(i, DOGSFSTAT_COL)    ' ������ �������� � SF
            ContrK = .Cells(i, 8)
            ContrId = ContractId(ContrK)
            If .Cells(i, 1) = 1 Then    ' � SF ���� ���.
                If ContrId <> "" Then   '==== ������� ������ � SF?
                    OppN = WorksheetFunction.VLookup(ContrK, _
                        Sheets(SFD).Range("B:Q"), 16, False)
                    If OppN = 0 Then
                        Call OppForDog(i)
                    Else
                '---- ������ ������������� ������� -------
                        OppIs = OppIs + 1
                        OppTyp = WorksheetFunction.VLookup(OppN, _
                            Sheets(SFopp).Range("B:R"), 17, False)
                        Vendor1C = .Cells(i, VENDOR1C_COL)
                        If Not IsSameVendor(OppTyp, Vendor1C, ContrK) And _
                                    Vendor1C <> "" Then
                            LogWr "     WARNING:  ���������1� = '" & Vendor1C & _
                                "' ���... = " & OppTyp & _
                                " � OppN = " & OppN & " ������������! ������� " _
                                & ContrK
                        End If
                    End If
                Else                    '===== ������ ����� �������
                    NewContract .Cells(i, DOG1C_COL), .Cells(i, DOG1C_MAINDOG_COL), ContrK
                End If
            Else
'!!                NewAcc .Cells(i, DOG1CACC_COL)
            End If
        End With
    Next i

    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Dogovor\"
    WriteCSV NewContract, "Dogovor.txt"
    WriteCSV NewContractLnk, "ContrUpd.txt"
    Shell "quotaDogovor.bat"
    Shell "quotaContUpd.bat"

    MsgBox "�������: " & vbCrLf & vbCrLf _
        & "   ������� ����������� " & Fruitful _
        & vbCrLf & "   ������� ����=" & OppIs _
        & vbCrLf & "  ����� ��������=" & OppNew & " ��������=" & NoOpp

    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\OppInsert\"
    WriteCSV O_NewOpp, "OppInsert.txt"
    Shell "quota2.bat"

    ModEnd DOG_SHEET
End Sub
Function OppForDog(iDog) As Boolean
'
' ���� �� ���������� ������ ��� ���������� � ��� ��������?
' ���� ���� - ��������� (ContrLink); ���� ��� - ������� ����� (NewOpp)
' �������� ������ ���������� � ������ ��������� DOG_SHEET (=3), iDog - ��� ������
'   22.3.12

    Dim OppLnk As String        ' ����� �������, � ������� ����� ������� �������
    Dim Acc, Dat, Sale, DogValue, DogCur As String  ' ��������� ������ �������
    Dim Stage, t As String      ' ���� � ��� ������ ������� �� ��������
    Dim ContrK As String        ' ��� �������� � ����� ���/�������
    
    Dim t0, Tend
    
    t0 = Timer
        
    With Sheets(DOG_SHEET)
        Acc = .Cells(iDog, DOG1CACC_COL)
        Dat = .Cells(iDog, DOG1CDAT_COL)
        Sale = .Cells(iDog, DOG1CSALE_COL)
        DogCur = CurISO(.Cells(iDog, DOGCUR1C_COL))
        DogValue = .Cells(iDog, DOGSUM1C_COL) * CurRate(DogCur)
        ContrK = .Cells(iDog, DOGCOD_COL)
        t = OppT(.Cells(iDog, VENDOR1C_COL), Sale)
           
        OppLnk = IsOpp(Sale, Acc, t, DogValue, Dat, ContrK)
    
        If OppLnk = "" Then     ' ���� ���������� ������, ���� ������ �������?
                                    '---- ��� - ������� ����� ������
            Stage = "70%-�������� �� ������� � ���������"
            If .Cells(iDog, DOGINVC_COL) <> 0 Then
                Stage = "80%-�������� �������� ��� ���� � ������"
            End If
            If .Cells(iDog, DOGPAID_COL) <> 0 Then
                Stage = "90%-������ ������ ������ �� ����"
            End If
'!!            NewOpp Acc, ContrK, Dat, Sale, DogValue, DogCur, T, Stage
        Else                        '---- ���� - ��������� � ��� �������
            Call ContrLink(ContrK, OppLnk)
            LogWr "<L> ����� �������� '" & ContrK & "' � �������� " & OppLnk
        End If
    End With
    Tend = (Timer - t0) * 1000
'    MsgBox "iDog = " & iDog & "Timer=" & Tend & " ����"
'    If iDog > 100 Then Stop
End Function
Function OppT(V1C, Sale) As String
'
' ���������� ��� ������� SF �� ����� ���������� � �������� 1�
'   3.2.12
'  20.3.12 - ���� ��������������� � ������� We "����&��� ������������"

    Dim S As Range
    OppT = V1C
    If OppT = "" Then Exit Function
    For Each S In Range("����_���_������������").Rows
        OppT = S.Cells(1, 1)
        If InStr(S.Cells(1, 2), V1C) <> 0 Then
            Select Case S.Cells(1, 3)
            Case "1:1", "1:M"
                Exit Function
            Case "M:1"
                If IsRightSale(Sale, OppT) Then Exit Function
                OppT = ""
            Case Else
                LogWr "FATAL ERROR OppT: �������� ������� W� ����&��� ������������"
                StopSub
            End Select
        End If
    Next S
End Function
Sub DogOppLink()
'
' [*] ������ �� SFD � ����� ���������� �������� ��� �����, Update Dogovor.csv
'   14.2.2012

    
    Dim i, j, Ndog As Integer
    Dim SaleDog, SaleOpp, OppN, OppType
    Dim DogDateStart, DogDateEnd, OppCloseDate As Date
    Dim Acc As String
    Dim ContrCode, Vendor1C As String

    Lines = ModStart(SFD, "DogOppLink: ����� �������� � ��������", True) - SFresLines

    CheckSheet SFD, Lines + 2, 3, SFcontrRepName
    CheckSheet SFopp, EOL_SFopp + 2, 1, SFoppRepName

    ClearSheet NewContractLnk, Range("HDR_ContrLnk")
    
'-- ������ �� ����� SFD - �� ���������
    For i = 2 To Lines
        Progress i / Lines
        With Sheets(SFD)
            If Trim(.Cells(i, 1)) <> "" Then    ' �������� ������ � ���������� ��� � SF
                Acc = .Cells(i, SFD_ACC1C_COL)
                SaleDog = .Cells(i, SFD_OWNER_COL)
                ContrCode = .Cells(i, SFD_COD_COL)
                Vendor1C = WorksheetFunction.VLookup(ContrCode, _
                    Sheets(DOG_SHEET).Range("H2:P77777"), 2, False)
                DogDateStart = .Cells(i, SFD_DATESTART_COL)
                DogDateEnd = .Cells(i, SFD_DATEEND_COL)
                If .Cells(i, SFD_OPPN_COL) = "" Then     '==== ������� ������ � ��������?
                '-- ������ �� SFopp - �� ��������           ��� - ���������
                    For j = 2 To EOL_SFopp
                        With Sheets(SFopp)
                            SaleOpp = .Cells(j, SFOPP_SALE_COL)
                            OppN = .Cells(j, SFOPP_OPPN_COL)
                        ' ��� ������� �������� �������� ������ ������� � ��� �� ������������
                        '       ..� � ���������� (��� � ��� ��) ���������
                            If .Cells(j, SFOPP_ACC1C_COL) = Acc And _
                                    IsSameTeam(SaleDog, SaleOpp, OppN) Then
                                OppType = .Cells(j, SFOPP_TYP_COL)
                                OppCloseDate = .Cells(j, SFOPP_CLOSEDATE_COL)
                        ' �������� ��  ���������� � ����� ��������
                                If IsSameVendor(OppType, Vendor1C, ContrCode) And _
                                        OppCloseDate <= DogDateEnd And _
                                        OppCloseDate >= DogDateStart Then
                                    Call ContrLink(ContrCode, .Cells(j, SFOPP_OPPID_COL))
                                End If
                            End If
                        End With
                    Next j
                Else                                    '==== ��� ������ � ��������
                    Dim DogPaidSum                          ' ����� �������� �� ��������
                    DogPaidSum = WorksheetFunction.SumIf(Sheets(SF).Range("Q:Q"), _
                        .Cells(i, SDF_CONTRN_COL), Sheets(SF).Range("F:F"))
                    If Deviation(DogPaidSum, .Cells(i, SFD_VALUE_COL)) < 0.05 Then
                        Dim DogStatus
                        DogStatus = .Cells(i, SFD_STATUS_COL)
                        If DogPaidSum = 0 Then
                            LogWr " ������� �� ������� � ��� ����� =0"
                        ElseIf DogStatus <> DOG_STAT_CLOSED Then
                            LogWr "<!> ������� " & .Cells(i, SFD_COD_COL) _
                                & " ( SFD � = " & .Cells(i, SDF_CONTRN_COL) _
                                & ") ���� �������"
                        End If
                    Else
                        i = i
                    End If
                End If
            End If
        End With
    Next i
    
'    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\OppInsert\"
'    WriteCSV O_NewOpp, "OppInsert.txt"
'    Shell "quota2.bat"
    ModEnd SFD
End Sub
Function Deviation(x, Y)
'
' ���������� ������������� ������� � � Y
'   15.3.12

    Const ErrVal = 999999
    If x <> 0 Then
        Deviation = Abs((x - Y) / x)
    ElseIf Y = 0 Then
        Deviation = 0
    Else
        Deviation = ErrVal
    End If
End Function
Function IsSameVendor(OppType, V1C, ContrCode) As Boolean
'
' ���������� True ���� ��� ������� ������������� ���������� � 1� �� ������� �� We
'   5.3.12
' 17.3.12 - ���������� LogWr

    Dim VendorSF As String  ' ���&��� ������������ �� �������
    
    Dim OppTypeArr() As String
    Dim i
    
    IsSameVendor = False
    
' ���� �� ����� �������� �������� � OppType -- � ���� �� ����� ���� ���������
    OppTypeArr = Split(OppType, ";")
    For i = 0 To UBound(OppTypeArr)
        VendorSF = ""
        On Error Resume Next
        VendorSF = WorksheetFunction.VLookup(Trim(OppTypeArr(i)), _
            Sheets(We).Range("����_���_������������"), 2, False)
        On Error GoTo 0
        If VendorSF = "" And V1C <> "" Then
            LogWr "     WARNING: [IsSameVendor] ����������� ���&��� = '" _
                & OppType & "' � SF, � � �������� 1� '" & ContrCode _
                & "' ��������� = '" & V1C & "'"
        End If
        If InStr(VendorSF, V1C) <> 0 Then GoTo Found
    Next i
    Exit Function
Found:
    IsSameVendor = True
End Function
