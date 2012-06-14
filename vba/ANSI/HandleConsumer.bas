Attribute VB_Name = "HandleConsumer"
'--------------------------------------�������� ��������----------------
' ConsumerHandling - ������ ��� ������ �� �����������
'   ������������ � ����������� ����������� ����� ������� �������
'   � ���� ����������������� ����� ������
'
' (*) ConsOppCreate - �������� �������� SF �� ������ Match � ������ ����� DL
' (*) PaidConsumers - ������ �������� �� ����������� ����� DL
'   14.2.2012

Option Explicit

Public Const NewOppSheet = "O_NewOpp"
Public Const Consumers = "����������"
Public Const ConumersActivity = "������� ��������� ���������� � ���"
Dim OppN As Integer
Sub ConsOppCreate()
'
' ������ �� ������ �������� � �������� "bulky" �������� �� �����������
'   12.2.2012

    Dim i As Integer
    Dim Acc, SalesRep As String
    
    Lines = ModStart(1, "������� �� �����������") - 3
    
    ClearSheet NewOppSheet, 4
    OppN = 1
'    Lines = 250
'------ ���������� ������� ----------------
    For i = 2 To Lines
        SalesRep = Sheets(1).Cells(i, 22)
        If SalesRep = "������" Or SalesRep = "������" Then
            Acc = Sheets(1).Cells(i, 9)
            If IsOpp(Acc, ConumersActivity) = 0 And _
                    Sheets(1).Cells(i, 1) = 1 And _
                    Sheets(1).Cells(i, 4) <> 1 Then
                NewOpp Acc, Acc & "-" & "����������", "1.1.2020", "������", _
                    500000, "RUB", "����������"
            End If
        End If
    Next i
    
    Columns("E:E").Select           ' ����������� ������ ���������� �������
    Selection.NumberFormat = "0"
    
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\OppInsert\"
    WriteCSV NewOppSheet, "OppInsert.txt"
    Shell "quota2.bat OppInsert.TXT C:\SFconstr\OppInsert.csv"

    ModEnd NewOppSheet
End Sub
Sub NewOpp(Account, OppName, CloseDate, OwnerOpp, Value, CurrencyOpp, TypOpp)
'
' ����� ������ ��� ������ DL � ����������� Account
'   12.2.2012

    Dim i, AccountId, OwId
    
    AccountId = AccId(Account)
    OwId = OwnerId(OwnerOpp)
    Sheets(NewOppSheet).Select
    
    For i = 1 To OppN
        If Cells(i, 1) = AccountId Then Exit Sub    ' ������������ NewOpp
    Next i
    
    OppN = OppN + 1
    Cells(OppN, 1) = AccountId
    Cells(OppN, 2) = OppName
    Cells(OppN, 3) = CloseDate
    Cells(OppN, 4) = OwId
    Cells(OppN, 5) = Value
    Cells(OppN, 6) = CurrencyOpp
    Cells(OppN, 7) = TypOpp
    If TypOpp = Consumers Then
        Cells(OppN, 8) = "90%-������ ������ ������ �� ����"
        Cells(OppN, 9) = ConumersActivity
        Cells(OppN, 10) = "��������� ��������� � ���"
    End If
End Sub


Function Adr1C(Acc) As String
'
'   ���������� ����� ����������� Acc � ������ �������� 1� ��� ������, ���� �� ���
'   13.2.2012
    On Error Resume Next
    Adr1C = 0
    Adr1C = WorksheetFunction.VLookup(Acc, Sheets(Acc1C).Range("E:H"), 4, False)
    On Error GoTo 0
End Function
Function AccId(Account)
'
' ��������� Id ����������� �� ������ SFacc
'   13.2.2012
    On Error Resume Next
    AccId = 0
    AccId = WorksheetFunction.VLookup(Account, Sheets(SFacc).Range("B:C"), 2, False)
    On Error GoTo 0
End Function
Function OwnerId(Owner)
'
' ��������� Id �� ������� �������� � �������������� �� ������� "We"
'   13.2.2012
    On Error Resume Next
    OwnerId = 0
    OwnerId = WorksheetFunction.VLookup(Owner, Sheets(We).Range("A8:C21"), 3, False)
    On Error GoTo 0
End Function
Function IsOpp(Account, T)
'
' ��������, ���� �� � ����������� Account ���������� ������ ���� �.
' ���� ���� - ���������� ����� ������ �������, ���� ��� - 0
'   13.2.2012

    Dim i
    
    IsOpp = 0
    If AccId(Account) = 0 Then Exit Function    ' ���� ��� ��� � SF -> =0
    For i = 2 To EOL(SFopp)
        If Cells(i, 4) = Account And Cells(i, 6) = 0 And Cells(i, 17) = T Then
            IsOpp = Cells(i, 2)                 ' ���� ������ ������� ���� � �� �� ������
            Exit For
        End If
    Next i
End Function

Sub PaidConsumers()
'
' ������ �� ������ �������� � ��������� �����������
'   14.2.2012

    Dim i, j, a, Pct
    Dim Account As String
    Const SheetN = "P_PaidConsumer"
    
    ClearSheet SheetN, 7
    Lines = ModStart(1, "��������� �������� �� �����������")
    
    j = 2
    For i = 2 To Lines
        Call Progress(i / Lines)
'        Sheets(1).Select
'        Cells(i, 1).Select
        Account = Sheets(1).Cells(i, 9)
            ' ���� ������ � ������ �� ������ � ��� ��� � SF
        If IsOpp(Account, ConumersActivity) <> 0 And _
            Sheets(1).Cells(i, 4) <> 1 And _
            Trim(Sheets(1).Cells(i, 6)) <> "" Then
                Sheets(SheetN).Cells(j, 1) = Sheets(1).Cells(i, 6)              ' ����.���.
                Sheets(SheetN).Cells(j, 2) = DDMMYYYY(Sheets(1).Cells(i, 7))    ' ����
                Sheets(SheetN).Cells(j, 3) = Sheets(1).Cells(i, 8)              ' ����
                Sheets(SheetN).Cells(j, 4) = Dec(Sheets(1).Cells(i, 18))        ' ���� ���.
                Sheets(SheetN).Cells(j, 5) = Sheets(1).Cells(i, 19)             ' �����
                Sheets(SheetN).Cells(j, 6) = ContractId( _
                    Sheets(1).Cells(i, 25), Sheets(1).Cells(i, 26))             ' ContractId
                Sheets(SheetN).Cells(j, 7) = IsOpp(Account, ConumersActivity)   ' OppN
             Sheets(SheetN).Select
             Cells(j, 1).Select
             j = j + 1
        End If
    Next i
    
' !!!! ���������, ��� �� ����� �������� ����� "�������" � SF
'    Lines = EOL(SFD)
'    For i = 2 To EOL(SheetN)
'
    
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\Payment\"
    WriteCSV SheetN, "Paid.txt"
    Shell "quota2.bat"

    ModEnd 1
End Sub
Function DDMMYYYY(d) As String
'
' �������������� ���� d � ��������� ������ DDMMYYYY
'   14.2.2012
    DDMMYYYY = Day(d) & "." & Month(d) & "." & Year(d)
End Function
Function Dec(a) As String
'
' �������������� ����� � � ��������� ������ � ���������� ������
'   14.2.2012

    Dec = "'" & WorksheetFunction.Substitute(a, ",", ".")
'    On Error Resume Next
'     If WorksheetFunction.Search(",", Dec) <> "#VALUE!" Then MsgBox "�������!"
'    On Error GoTo 0
End Function
Function ContractId(Contr, Optional MainContr = "") As String
'
' ��������� Id �� ����� ��������� � ��������� ���������, ���� ����
' ���� ��������� ��� � SF, ���������� 0
'   14.2.2012

    Dim C As String
    
    C = MainContr
    If C = "" Then
        C = Contr
    Else
        C = C & "/" & Contr
    End If
    On Error Resume Next
    ContractId = ""
    ContractId = WorksheetFunction.VLookup(C, Sheets(SFD).Range("2:15"), 13, False)
    On Error GoTo 0
End Function
