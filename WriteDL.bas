Attribute VB_Name = "WriteDL"
'================================================================================
' ������ ��� �������� ������ ����������� Informatica Data Loader
'
' * NewContractDL  - � ������ 1� �������� ����� ������� - ������ � Dogovor.csv
' - PaymentUpd(PayKod, ContrId) - Update ������� - ������ � SF ����� Data Loader
' - NewContract(Dogovor, MainDog, ContrK) - �������� ������ �������� ContrK � SF
' - DogFormat(Wsheet) - �������������� ����� Wsheet ��� ������ � Dogovor.csv
' S WrCSV(SheetN, Directory, FileToWrite, Bat) - ��� - ������ CSV
' - WriteCSV(SheetN, FileName, ..)  - ������ ����� SheetN � ���� ��� �������� DL
'
'   23.9.2012 - ��������� ������ AddressParse
'   31.10.12

Option Explicit

Const ACC_NEWDOG_COL = 7    ' ������� ������

Const C1C = 3       ' ����� ����� ������ �� ��������� �� 1�
Dim C1Cname         ' ��� ����� ���������, ��� �������� ����

Sub NewContractDL()
'
' ���������� ����� Dogovor.csv �� ��� DL Task "Dogovor_Insert"
'   24.1.2012

    Const ResLines = 7      ' ����� ����� ������ ��� ������� �� ��������� (�����)

    Dim MyRange As Range
    Dim W, L, Wsheet
    
    EOL_SFD = EOL(SFD) - SFresLines
    
    Sheets(3).Select
    W = ActiveSheet.UsedRange.Columns.Count
    L = ActiveSheet.UsedRange.Rows.Count - ResLines
    
    Sheets(3).Select
    Set MyRange = Range(Cells(1, 1), Cells(L, W))
    MyRange.AutoFilter Field:=1, Criteria1:="<>"            ' ����������� ���� � SF
    MyRange.AutoFilter Field:=2, Criteria1:="=��� � SF"     ' �������� ��� � SF
    Range(Cells(1, 10), Cells(L, 23)).Copy
    
    Sheets.Add After:=Sheets(Sheets.Count)  ' ������� ����� ���� - ���������
    Wsheet = ActiveSheet.Name
    ActiveSheet.Paste
    
    Call DogFormat(Wsheet)  ' ����������� ������� ���� Wsheet ��� ������ ��������
    
    ChDir "C:\Users\salesforce\Dogovor"
    WriteCSV Wsheet, "Dogovor.txt"

    Shell "quota2.bat"
    ActiveSheet.Delete                      ' ���������� ��������� ����
End Sub
Sub PaymentUpd(PayKod, ContrId)
'
' Update ������� - ������ � SF ����� Data Loader
' �� ����������� ���� ��������� ������ ���� - ������ �� �������
'   11.3.12

    Dim PaymentId As String

    PaymentId = ""
    On Error Resume Next
    PaymentId = WorksheetFunction.VLookup(PayKod, _
        Sheets(SF).Range("B:R"), 17, False)
    On Error GoTo 0
    If PaymentId = "" Then GoTo Err
    EOL_PaymentUpd = EOL_PaymentUpd + 1
    Sheets(P_PaymentUpd).Cells(EOL_PaymentUpd, 1) = PaymentId
    Sheets(P_PaymentUpd).Cells(EOL_PaymentUpd, 2) = ContrId
    Exit Sub
Err:
    MsgBox "�� ������ ������ SF �� ���� '" & PayKod & "'", , "FATAL ERROR!"
    Stop
End Sub
Sub NewContractXX(Dogovor, MainDog, ContrK)
'
' XXX �������� ������ �������� ContrK � SF
'   31.3.12

    Dim StartDate, OurLegalName, IdWe, Salesman, IdOwner As String
    Dim Client, DogType, Value, DogCurrency, CurCod, Dat As String

    Dim i
       
' ������� ������ �������� �� ������ 1�
    With Sheets(DOG_SHEET)
        For i = 2 To EOL_DogSheet
            If ContrK = .Cells(i, DOGCOD_COL) Then
                StartDate = .Cells(i, DOG1CDAT_COL)
                OurLegalName = .Cells(i, DOG1CFIRM_COL)
                IdWe = WorksheetFunction.VLookup(OurLegalName, Range("��"), 2, False)
                Salesman = .Cells(i, DOG1CSALE_COL)
                IdOwner = ""
                On Error Resume Next
                IdOwner = WorksheetFunction.VLookup(Salesman, Range("��������_�_SF"), 3, False)
                On Error GoTo 0
                If IdOwner = "" Then IdOwner = "00520000001460C"    ' ��� ���
                Client = .Cells(i, DOG1CACC_COL)
                DogType = .Cells(i, DOGTYPE_COL)
                Value = Dec(.Cells(i, DOGSUM1C_COL))
                DogCurrency = .Cells(i, DOGCUR1C_COL)
                Dat = DDMMYYYY(.Cells(i, DOG1CDAT_COL))
                Exit For
            End If
        Next i
    End With
    
' ����� Id ��������� �������� � SFD
    Dim IdMainDog As String
    Dim X() As String
    
    IdMainDog = ""
    If MainDog <> "" Then
        X = split(Compressor(MainDog), " ")
        On Error Resume Next
        IdMainDog = WorksheetFunction.VLookup(X(1), _
            Sheets(SFD).Range("B:O"), 14, False)
        On Error GoTo 0
        If IdMainDog = "" Then
            LogWr "NewContract: �� ������ �������� ������� �� '" _
                & ContrK & "' -- ������� � SF �� �������!"
            Exit Sub
        End If
    End If

' ���������� ����� ������ ��� ����� �� SFacc
    Dim Street, City, State, Index, Country As String
    With DB_SFDC.Sheets(SFacc)
        For i = 2 To EOL_SFacc
            If .Cells(i, SFACC_ACC1C_COL) = Client Then
                Street = .Cells(i, SFACC_STREET_COL)
                City = .Cells(i, SFACC_CITY_COL)
                State = .Cells(i, SFACC_STATE_COL)
                Index = .Cells(i, SFACC_INDEX_COL)
                Country = .Cells(i, SFACC_COUNTRY_COL)
                Exit For
            End If
        Next i
    End With
    
' ���������� ����� NewContract ������� �� ������ ��������� ��� CSV
    EOL_NewContr = EOL_NewContr + 1
    With DB_MATCH.Sheets(NewContract)
        .Cells(EOL_NewContr, NEWDOG_DOGOVOR_COL) = " " & Dogovor
        .Cells(EOL_NewContr, NEWDOG_DATE_COL) = DDMMYYYY(StartDate)
        .Cells(EOL_NewContr, NEWDOG_WE_COL) = OurLegalName
        .Cells(EOL_NewContr, NEWDOG_IDWE_COL) = IdWe
        .Cells(EOL_NewContr, NEWDOG_SALESMAN_COL) = Salesman
        .Cells(EOL_NewContr, NEWDOG_IDOWNER_COL) = IdOwner
        .Cells(EOL_NewContr, NEWDOG_CLIENT_COL) = Client
        .Cells(EOL_NewContr, NEWDOG_IDACC_COL) = AccId(Client)
        .Cells(EOL_NewContr, NEWDOG_TYPE_COL) = DogType
        .Cells(EOL_NewContr, NEWDOG_MAINDOG_COL) = IdMainDog
        .Cells(EOL_NewContr, NEWDOG_VALUE_COL) = Value
        .Cells(EOL_NewContr, NEWDOG_CURRENCY_COL) = DogCurrency
        CurCod = CurISO(DogCurrency)
        .Cells(EOL_NewContr, NEWDOG_CURRENCYCOD_COL) = CurCod
        .Cells(EOL_NewContr, NEWDOG_RATE_COL) = Dec(CurRate(CurCod))
        .Cells(EOL_NewContr, NEWDOG_RATEDATE_COL) = Dat
        .Cells(EOL_NewContr, NEWDOG_ACCSTREET_COL) = Street
        .Cells(EOL_NewContr, NEWDOG_ACCCITY_COL) = City
        .Cells(EOL_NewContr, NEWDOG_ACCSTATE_COL) = State
        .Cells(EOL_NewContr, NEWDOG_ACCINDEX_COL) = Index
        .Cells(EOL_NewContr, NEWDOG_ACCCOUNTRY_COL) = Country
    End With
    Exit Sub
Err:    MsgBox "NewContract '" & ContrK & "'", , "ERROR!"
    Stop
End Sub
Sub ContrLink(ContrCode, OppId)
'
' ���������� �������� (��� ContrK) � �������� (OppId)
'   18.3.12


    EOL_ContrLnk = EOL_ContrLnk + 1
    With Sheets(NewContractLnk)
        .Cells(EOL_ContrLnk, 1) = ContractId(ContrCode)
        .Cells(EOL_ContrLnk, 2) = OppId
    End With
End Sub
Sub DogFormat(Wsheet)
'
' ������������ �������������� �������� ����� Wsheet ��� ������ � Dogovor.csv
'   25.1.2012
'   7.3.12 - ������� ������ ������� � ����� �������
'   14.3.12 - ������ � �������� ���������
    
    Const DogHdr = "DogovorHeader"  ' ������ �������� ��� DL Dogovor_Insert
    Dim MyRange As Range
    Dim L, i
    
    Sheets(DogHdr).Rows("1:3").Copy
    Sheets(Wsheet).Select
    Lines = ActiveSheet.UsedRange.Rows.Count
    L = Lines + 3
    Set MyRange = Range("A1:U" & L)
    
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown     ' ������ �� �������
            
    Columns("B:B").NumberFormat = "@"   ' �������� ���� �� �����
    For i = 5 To L
        Cells(i, 2) = Day(Cells(i, 3)) & "-" & Month(Cells(i, 3)) & "-" & Year(Cells(i, 3))
    Next i
    Range("C4:C" & L).Delete Shift:=xlToLeft     ' B - Delete - ���� ���������� ��������
    
 ' ��������� ������� D - IdWe � ��������� �� �� ������� VLOOKUP
    Range("D4:D" & L).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D3").AutoFill Destination:=Range("D3:D" & L), Type:=xlFillDefault
    
    For i = 5 To L
        Cells(i, 6) = Compressor(Cells(i, 6)) ' ������� ������ ������� � ����� �������
    Next i

 ' ��������� ������� F - IdOwner � ��������� F � H ��������� �� �������
    Range("F4:F" & L).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F3").AutoFill Destination:=Range("F3:F" & L), Type:=xlFillDefault
    Range("H3").AutoFill Destination:=Range("H3:H" & L), Type:=xlFillDefault

    Columns("L:L").NumberFormat = "@"   ' �������� ���������� ������� �� ����� � �����
    Call MainContrIdSet(Wsheet, L)      ' �������� ���� �������� �������� �� Id �� SFD
    
    For i = 5 To L
        Cells(i, 12) = Replace(Cells(i, 12), ",", ".")
    Next i

    Range("M4:N" & L).Delete Shift:=xlToLeft    ' ������� ������� ���� �� ������ � ��������

    Range("N3:U" & L).FillDown          ' ��������� ������� �� ������� �� ��������� �����
    Rows("2:4").Delete                  ' ������� ������ �� �������
    Rows("1:1").AutoFilter
    MyRange.AutoFilter Field:=4, Criteria1:="<>#N/A"
    MyRange.AutoFilter Field:=6, Criteria1:="<>#N/A"
    MyRange.AutoFilter Field:=8, Criteria1:="<>#N/A"
    MyRange.AutoFilter Field:=14, Criteria1:="<>#N/A"
    
    Range("A1:U" & Lines).Copy Cells(Lines + 2, 1)  ' �������� ��������, � ������� ��� #N/A
    Call AutoFilterReset(Wsheet)
    Rows("1:" & Lines + 1).Delete

    MsgBox "�� ��������� ������������ ���������" & vbCrLf & _
            "     �� ������ ��� ����� SF - SFD!", vbExclamation, "��������!"
End Sub
Sub MainContrIdSet(Wsheet, L)
'
' ����������� Id ��������� ��������� ��� �� ������ � ������ �� ��������� 1� �� SFD
'   14.3.12

    Const MAIN_DOG_SFD_COL = 5      ' ������� "���.�������"
    Const ACC_SFD_COL = 6           ' ������� "��� ����������� 1�" � SFD
    Const MAIN_DOGID_SFD_COL = 19   ' ������� "��� ��������� ��������"
    
    Dim i, j, IsMain, X
    Dim MainContr, MainContrId, Account As String
    
    For i = 5 To L
        MainContr = Sheets(Wsheet).Cells(i, 10)
        If MainContr <> "" Then
            MainContrId = ""
            Account = Sheets(Wsheet).Cells(i, ACC_NEWDOG_COL)
            With Sheets(SFD)
                For j = 2 To EOL_SFD
                    X = .Cells(j, MAIN_DOG_SFD_COL)
                    IsMain = InStr(MainContr, X)
                    If IsMain <> 0 And X <> "" Then
                        If .Cells(j, ACC_SFD_COL) = Account Then
                            MainContrId = .Cells(j, MAIN_DOGID_SFD_COL)
                            GoTo FoundMainContr
                        End If
                    End If
                Next j
            End With
            LogWr "     WARNING:  �� ������ � SFD �������� ������� '" _
                & MainContr & "'"
            MainContrId = "�� ������ �������� �������!"
FoundMainContr:
            Sheets(Wsheet).Cells(i, 10) = MainContrId
        End If
    Next i
End Sub
Sub WrCSV(SheetN As String, Directory, FileToWrite, Bat)
'
' S WrCSV(SheetN, Directory, FileToWrite, Bat) - ��� - ������ CSV
'   23.9.12
'   30.9.12 - ������ ����� ����� ����� � SheetN � WrProcResult
'   19.10.12 - ������������� WP_TMP
'   31.10/12 - �� ������� ���� CSV

    Dim NewLines As Long
    
    Set DB_TMP = FileOpen(F_TMP)
    DB_TMP.Sheets(SheetN).Activate
    NewLines = EOL(SheetN, DB_TMP)
    If NewLines <= 1 Then
'        Application.DisplayAlerts = False
'        ActiveSheet.Delete
'        Application.DisplayAlerts = True
    Else
        ChDir Directory
        WriteCSV SheetN, FileToWrite
        Shell Bat
        LogWr "WrCSV> � '" & FileToWrite & "' �������� " & NewLines & " �����."
    End If
    
    WrProcResult NewLines - 1
End Sub
Sub WriteCSV(SheetN, FileName, _
    Optional Row0 As Integer = 1, Optional Col0 As Integer = 1, _
    Optional RowLast As Integer, Optional ColLast As Long)
'
' ������ ����� SheetN � ���� FileName � ������� CSV Tab Delimited � ��� ��������
'   24.1.2012
'    3.2.2012 Optionan Range ������������� �������
'   13.2.2012 EOL(SheetN), ����� �� ������ ������ ������ � �����
'   29.2.2012 bug fix - Cells ��� Sheets �������� � ������ ������������ ������
'   23.9.12 - ������� � match 2.0
    
    Dim S, Col
    
    Open FileName For Output As #1
    
    If RowLast = 0 Then RowLast = EOL(SheetN)
    If ColLast = 0 Then
        ColLast = ActiveSheet.UsedRange.Columns.Count
    End If
    For S = Row0 To RowLast
        For Col = Col0 To ColLast
            Print #1, ActiveSheet.Cells(S, Col);
            If Col <> ColLast Then Print #1, vbTab;
        Next Col
        Print #1, " "
    Next S
    Close #1
'    MsgBox "�������� " & (RowLast - Row0 + 1) & " �����" & vbCrLf & "� ���� " & Chr(171) & FileName & Chr(187)
End Sub

Sub WritePaid(FileName)
'
' ������������ ������ ���� FileName ��� ��������� DL Payment
'   10.1.2012
    
    Range("C1:I" & ActiveSheet.UsedRange.Rows.Count).Copy   ' �������� ������ ������� �� P_Update
    Sheets.Add After:=Sheets(Sheets.Count)  ' ������� ����� ���� - ���������
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False           ' �������� � ���� �� P_Update Paste Value

    Open FileName For Output As #1
    
    With ActiveSheet.UsedRange
        For i = 1 To .Rows.Count
            Print #1, Range("A" & i) & vbTab & Range("B" & i) & vbTab & Range("C" & i) & vbTab & _
                Range("D" & i) & vbTab & Range("E" & i) & vbTab & Range("F" & i) & vbTab & Range("G" & i)
        Next
    End With
    Close #1
    ActiveSheet.Delete                      ' ���������� ��������� ����
End Sub



