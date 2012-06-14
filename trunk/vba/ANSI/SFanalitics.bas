Attribute VB_Name = "SFanalitics"
'--------------------------------------------------------------------
' SFanalitics  - ������������ ��������� �������� Salesforce.com:
'       ����������� = Accounts      = Acc   ����� SFacc
'       ��������    = Opportunity   = Opp   ����� SFopp ��� SF
'       �������� -- ��� ���������� � SF --  ����� SF
'       ���������   = Contract      = Contr ����� SFD
'   24.5.2012
'
' - AccId(Account)              - Id SF ����������� �� ����� 1�
' - OwnerId(Owner, Buddy)       - Id SF ��������� � �������������� �� ������� We
' - OppByPay(PayKod)            - �������� ��� ������� SF �� ������� 1�
' - OppNbyPay(PayKod)           - �������� ����� ������� SF �� ������� 1�
' - OppNbyId(OppId)             - ��������� ������ ������� OppN �� ��� Id � SF
' - OppIdbyPay(PayKod)          - �������� Id SF ������� �� ������� 1�
' - OppOwner(PayKod)            - ���������� ��������� ������� SF �� ���� ������� 1�
' - OppNameById(OppId)          - ��������� ����� ������� �� ��� Id � SPopp
'??' - IsOpp(Sale, Account, T)     - ���������, ���� �� ������������� ������ ���� �
' - ContrNbyPay(PayKod)         - ��������� ������ �������� �� ���� ������� �� ����� SF
' - ContractId(ContrCod)        - ��������� Id SF �� ��������e � 1�
' - ContrCod(Contr, MainContr)  - ������������ ������ <���.�������/�������>
' - ContrCodByOppId(OppId)      - ���������� ��� �������� (���/�������) �� Id ������� � SF
' - ContrOK(OppN, ContrK)       - True ���� � SF ������ ������ � �������� � ��������� ��� � 1�
' - PayIdByK(PayK)  - ��������� Id SF �� ���� �������

Option Explicit

Function Adr1c(Acc) As String
'
'   ���������� ����� ����������� Acc � ������ �������� 1� ��� ������, ���� �� ���
'   13.2.2012
    On Error Resume Next
    Adr1c = 0
    Adr1c = WorksheetFunction.VLookup(Acc, Sheets(Acc1C).Range("E:H"), 4, False)
    On Error GoTo 0
End Function
Function AccId(Account)
'
' ��������� Id SF ����������� �� ������ SFacc �� ����� � 1�
'   13.2.2012
    On Error Resume Next
    AccId = 0
    AccId = WorksheetFunction.VLookup(Account, _
        Sheets(SFacc).Range("B:C"), 2, False)
    On Error GoTo 0
End Function
Function OwnerId(Owner, Buddy) As String
'
' ��������� SF OwnerId �� ������� �������� � �������� � ��������� 1C
'   3.3.2012
'   6.3.12 - ������� � ������� �������� Owner � We
'   12.3.12 - ����������� ���������� OppBuddy
    
    Dim x As Range

    For Each x In Range("��������").Rows
        If InStr(Owner, x.Cells(1, 1)) <> 0 Then GoTo Found
    Next x
    LogWr "ERROR! �� ������ �������� " & Owner & _
        ". ��� ������ � Match/We ��� ����� ��������� � 1�"
    Stop
Found:
'    If InStr(Owner, "������") Then
'''        MsgBox "������"
'    End If
    If InStr(x.Cells(1, 3), x.Cells(1, 1)) = 0 Then
        Buddy = x.Cells(1, 1)
    Else
        Buddy = ""
    End If
    OwnerId = x.Cells(1, 4)
End Function
Function OppByPay(PayKod)
'
' ��������� ����� ������� �� ���� ������� 1C �� ����� SF
'   17.2.2012
    On Error Resume Next
    OppByPay = ""
    OppByPay = WorksheetFunction.VLookup(PayKod, _
        Sheets(SF).Range("B2:K77777"), 10, False)
    On Error GoTo 0
End Function
Function OppNbyId(OppId) As Integer
'
' ��������� ������ ������� OppN �� ��� Id � SF
'   21.3.2012
'   25.4.12 - Lookup case insensitive - ������� �� ���� �� SFopp
    
    Dim i
    OppNbyId = 0
    For i = 2 To EOL_SFopp
        If Sheets(SFopp).Cells(i, SFOPP_OPPID_COL) = OppId Then
            OppNbyId = Sheets(SFopp).Cells(i, SFOPP_OPPN_COL)
            Exit For
        End If
    Next i
End Function

Function OppNbyPay(PayKod) As Integer
'
' ��������� ������ ������� �� ���� ������� 1C �� ����� SF
'   25.2.2012
    OppNbyPay = 0
    On Error Resume Next
    OppNbyPay = WorksheetFunction.VLookup(PayKod, _
        Sheets(SF).Range("B2:L77777"), 11, False)
    On Error GoTo 0
End Function
Function OppIdbyPay(PayKod)
'
' ��������� Id ������� �� ���� ������� 1C �� ����� SF
'   11.3.2012
    OppIdbyPay = ""
    On Error Resume Next
    OppIdbyPay = WorksheetFunction.VLookup(PayKod, _
        Sheets(SF).Range("B:S"), 18, False)
    On Error GoTo 0
End Function
Function OppIdbyContrId(ContrId)
'
' ��������� Id ������� �� ���� ������� 1C �� ����� SF
'   11.3.2012
    OppIdbyPay = ""
    On Error Resume Next
    OppIdbyPay = WorksheetFunction.VLookup(PayKod, _
        Sheets(SF).Range("B:S"), 18, False)
    On Error GoTo 0
End Function
Function OppOwner(PayKod)
'
' ��������� ������� �������� �� ���� ������� �� ����� SF
'   15.2.2012
    On Error Resume Next
    OppOwner = ""
    OppOwner = WorksheetFunction.VLookup(PayKod, Sheets(SF).Range("B2:I77777"), 8, False)
    On Error GoTo 0
End Function
Function OppNameByN(OppN)
'
' ��������� �������� ������� �� ����� SFopp �� ��� Id
'   23.3.2012
    Dim S As String
    S = OppN
    On Error Resume Next
    OppNameByN = ""
    OppNameByN = WorksheetFunction.VLookup(S, Sheets(SFopp).Range("B:C"), 2, False)
    On Error GoTo 0
End Function

Function ContrNbyPay(PayKod)
'
' ��������� ������ �������� �� ���� ������� 1C �� ����� SF
'   10.3.2012
    ContrNbyPay = ""
    On Error Resume Next
    ContrNbyPay = WorksheetFunction.VLookup(PayKod, _
        Sheets(SF).Range("B:Q"), 16, False)
    On Error GoTo 0
End Function
Function ContrN(ContrCod)
'
' ��������� ������ �������� �� ���� �������� (/) �� ����� SFD
'   10.3.2012
    ContrN = ""
    On Error Resume Next
    ContrN = WorksheetFunction.VLookup(ContrCod, _
        Sheets(SFD).Range("B:C"), 2, False)
    On Error GoTo 0
End Function
Function ContractId(ContrCod) As String
'
' ��������� Id �� ����� ��������� � ����� <���.�������/�������>
' ���� ��������� ��� � SF, ���������� ""
'   14.2.2012
'   28.2.2012 - ��������� ������������ - ���������� <���.�������/�������>

    ContractId = ""
    On Error Resume Next
    ContractId = WorksheetFunction.VLookup(ContrCod, _
        Sheets(SFD).Range("B:O"), 14, False)
    On Error GoTo 0
End Function
Function ContrCod(Contr, MainContr) As String
'
' �������������� � ����� <���.�������/�������>
'   28.2.12

    Dim C As String
    
    Contr = Trim(Contr)
    MainContr = Trim(MainContr)
    
    C = MainContr
    If C = "" Then
        C = Contr
    Else
        C = C & "/" & Contr
    End If
    ContrCod = C
End Function
Function ContrOppN(ContrId) As Integer
'
' ��������, ������ �� ������� ContrId � ��������. ���� �� - ���������� ����� �������
'   28.2.12
'   23.3.12 - ������-�� �� �������� Lookup - ������� �� ���� �� SFD
    
    Dim i
    ContrOppN = 0
    For i = 2 To EOL_SFD
        If Sheets(SFD).Cells(i, SFD_CONTRID_COL) = ContrId Then
            ContrOppN = Sheets(SFD).Cells(i, SFD_OPPN_COL)
            Exit For
        End If
    Next i
End Function
Function PayIdByK(PayK) As String
'
' - PayIdByK(PayK)  - ��������� Id SF �� ���� �������
'   13.5.2012

    PayIdByK = ""
    On Error Resume Next
    PayIdByK = WorksheetFunction.VLookup(PayK, _
        Sheets(SF).Range("B:R"), 17, False)
    On Error GoTo 0
End Function
Sub testPayInvByK()
    Call PayInvByK("542 �� 16.09.09��� ""������ ���""218")
    Call PayInvByK("542 �� 16.09.09���")                    ' �� ������!
    Call PayInvByK("549 �� 31.01.2012��� ""����""63")
End Sub
Function PayInvByK(PayK) As Integer
'
' - PayInvByK(PayK)  - ��������� ���� ����� ����� �� ���� �������
'   13.5.2012
'   24.5.12 - �������� ���� �� ���� ��-12 �� 01.02.11
    
    On Error Resume Next
    PayInvByK = InvoiceN(WorksheetFunction.VLookup(PayK, _
        Sheets(SF).Range("B:H"), 7, False))
    On Error GoTo 0
End Function
Function InvoiceN(Str) As Integer
'
' - InvoiceN(Str) ���������� ����� - ����� ����� �� ��������� ������ Str
'   24.5.12

    Dim x As Integer        ' ����� ������� - ������� � Str
    Dim y As Integer        ' ������� "C�-" � Str
    Dim S As String         ' �������� Str � ������� �����
    
    InvoiceN = 0
    If Str = "" Then Exit Function
    x = WorksheetFunction.Search(" ", Str)
    y = WorksheetFunction.Search("��-", Str)
    S = Mid(Str, y + 3, x - y - 3)
    
    If Not IsNumeric(S) Then Exit Function
    InvoiceN = S
End Function
Sub testInvoiceN()
    Call InvoiceN("""��-102 �� 28.02.11 ������ �� -�����""")
End Sub
Function IsRightContrOppLink(OppId, ContrCode)
'
' ���������� True, ���� ������ OppId ������ � ��������� ContrCode
'   1.4.12

    Dim i
    
    IsRightContrOppLink = False
    If OppId = "" Or ContrCode = "" Then Exit Function
    
    With Sheets(SFD)
        For i = 2 To EOL_SFD
            If .Cells(i, SFD_OPPID_COL) = OppId Then
                If .Cells(i, SFD_COD_COL) = ContrCode Then
                    IsRightContrOppLink = True
                    Exit Function
                End If
            End If
        Next i
    End With
End Function
Function ContrOK(OppN, ContrK) As Boolean
'
' ���������� True ���� � SF ������ ������������� ������ � �������� � ��������� ��� � 1�
'   8.3.12

    Dim DogOppN As Integer
    
    ContrOK = True
    If ContrK = "" Then Exit Function
    DogOppN = 0
    On Error Resume Next
    DogOppN = WorksheetFunction.VLookup(ContrK, _
        Sheets(SFD).Range("B:Q"), 16, False)
    On Error GoTo 0
    
    If DogOppN <> OppN Then ContrOK = False
End Function
