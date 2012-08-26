Attribute VB_Name = "From1C"
'---------------------------------------------------------------------------
' ������� ��� �������� ������� �� 1�
'
'<*> From1Cpayment  - �������� ���� ������ �� 1� "������ ����� �� �����"
' -  SFmatchFill(SheetN)  - ���������� ������ ����� SheetN �� SFDC
' -  CSmatch(Val,Col,[SheetN],[DB]) - Case Sensitive match ���������� ����� ������
'           � Val � ������� Col ����� SheetN � DB. ���� Val �� ������- ���������� 0.
'<*> From1Cdogovor  - �������� ���� ������ 1� "��������" ����� �� 1�
'<*> From1Caccount  - �������� ���� ������ 1� "������� .." ����� �� 1�
'(*) FromStock      - ������ ��������� ����� � ����� Stock
'
' 10.8.2012 �.�.������� match 2.0

Option Explicit
Sub SFlnkFill(DocFr, ColFr, ColFrId, ColVal, ColTo)
'
' - SFlnkFill(RepFr,ColFr,ColFrId, ColVal, ColTo) - "�������" ����� � DocFr � DocTo,
'       ��������� � ������� ColTo Id ������� ���������� �� �������� � ColFr
' 8.8.12

    Dim DocTo As String ' ��� �������� ��������� - ������
    Dim RepFr As TOCmatch, RepTo As TOCmatch
    Dim Val
    Dim i As Integer, N As Integer
    
    DocTo = ActiveSheet.Name
    RepTo = GetRep(DocTo)
    Workbooks(RepTo.RepFile).Sheets(RepTo.SheetN).Activate
    RepFr = GetRep(DocFr)
    With Workbooks(RepTo.RepFile).Sheets(RepTo.SheetN)
        Workbooks(RepFr.RepFile).Sheets(RepFr.SheetN).Select
        For i = 2 To RepTo.EOL
            Progress i / RepTo.EOL
            Val = .Cells(i, ColVal)
            N = CSmatch(Val, ColFr)
            If N > 0 Then
                .Cells(i, ColTo) = Workbooks(RepFr.RepFile).Sheets(RepFr.SheetN).Cells(N, ColFrId)
            Else
                .Cells(i, ColTo) = ""
            End If
        Next i
    End With
End Sub
Sub DogStatus()
'
' - DogStatus()     - ��������� � ���� ��������� ������ �� SND
'   11.8.12

    Dim i As Long, N As Long
    
    PublicProcessName = ""
    GetRep SFD
    GetRep DOG_SHEET
    DB_SFDC.Sheets(SFD).Activate
    With DB_1C.Sheets(DOG_SHEET)
        For i = 2 To RepTOC.EOL
            Progress i / RepTOC.EOL
            N = CSmatch(.Cells(i, DOGCOD_COL), SFD_COD_COL)
            If N > 0 Then
                .Cells(i, DOGSFSTAT_COL) = _
                    DB_SFDC.Sheets(SFD).Cells(N, SFD_STATUS_COL)
            End If
        Next i
    End With
End Sub
Sub ContractPaint()
'
' - ContractPaint() - ������������ ���� ���������
' 10.8.12

    GetRep DOG_SHEET
    Call AutoFilterReset(DOG_SHEET)
    Rows(1).RowHeight = 50
    
    Pnt DOGSFSTAT_COL, "������", rgbLightGreen      ' �������� �������� � SF- �������
    Pnt DOGSFSTAT_COL, "������", rgbOrange          ' �������� �������� - ���������
    Pnt DOGSFSTAT_COL, "��������", rgbLightBlue     ' ��������� - �������
    Pnt DOGSFSTAT_COL, "�� ���������", Antique      ' �� ��������� - Antique
    Pnt DOGSFSTAT_COL, "��� � SF", rgbWhite         ' ��� � SF - �� ����������
    Pnt DOGPAID1C_COL, 1, LimeG, 1                  ' ���������� - ����� �������
    Pnt DOGISINV1C_COL, 1, rgbOlive, 1              ' ��������� ���� - ���������
    Pnt DOG1CSCAN_COL, 1, rgbViolet, 1              ' ������������� - ����������
'-- �������� ����� � �������1�
    Range("Contract_Summary").Copy Destination:=ActiveSheet.Cells(RepTOC.EOL + 1, 1)
End Sub
Sub Acc1C_Bottom()
'
' - Acc1C_Bottom() - ������� �����x ���� ����� Acc1� � �����
'   14.8.12

    Dim R As TOCmatch
    Dim b As Range
    
    R = GetRep(Acc1C)
    DB_1C.Sheets(Acc1C).Activate
    Set b = ActiveSheet.Rows("1:3")
    b.Copy Destination:=Cells(R.EOL + 2, 1)
    b.Delete
End Sub
Sub AccPaint()
'
' - AccPaid() - ������� ������� � - ����������� ���� � SF
'   14.8.12

    Dim i As Long
    Dim RepTo As TOCmatch
    Dim R As Range
    
    RepTo = GetRep(ActiveSheet.Name)
    DB_1C.Sheets(RepTo.SheetN).Activate
    With Workbooks(RepTo.RepFile).Sheets(RepTo.SheetN)
        For i = 2 To RepTo.EOL
            Progress i / RepTo.EOL
            Set R = .Cells(i, PAYISACC_COL)
            If R <> "" Then
                R.Interior.Color = rgbYellow
            Else
                R.Interior.Color = rgbRed
            End If
        Next i
    End With
End Sub

Sub testCSmatch()
    If "G" = "g" Then Stop
    Dim A
    ThisWorkbook.Sheets("Sheet1").Select
    A = CSmatch("g12", 1)
    A = CSmatch("g121", 1)
    
    ModStart REP_1C_P_PAINT
    Set DB_SFDC = Workbooks.Open(F_SFDC, UpdateLinks:=False, ReadOnly:=True)
    DB_SFDC.Sheets(SFacc).Select
    A = CSmatch("��� ""���""", 2)
    ModEnd
End Sub
Sub PaymentPaint()
'
' - PaymentPaint() - ������������ ���� �������� 1C
' 24.6.12 ���������� ��� match 2.0
'  7.8.12 ��������� ��� ���

    Dim i As Integer
    Dim Rub, Doc    '���� "����� ���" � "����.���"
    
    PublicStepName = ""
    GetRep PAY_SHEET
    DB_1C.Sheets(PAY_SHEET).Select
    
    Range("A1:AC" & RepTOC.EOL).Interior.Color = rgbWhite   ' ���������� �������
    Rows("2:" & RepTOC.EOL).RowHeight = 15    ' ������ ����� �� ����� = 15
    
    With Sheets(PAY_SHEET)
        For i = 2 To RepTOC.EOL
            Progress i / RepTOC.EOL
            If .Cells(i, PAYINSF_COL) = 1 Then          ' ������� ������� � SF
                Range(Cells(i, 2), Cells(i, AllCol)).Interior.Color = rgbLightGreen
            ElseIf Trim(.Cells(i, PAYDOC_COL)) = "" Or Trim(.Cells(i, PAYSALE_COL)) = "" Then
                .Cells(i, 1).EntireRow.Hidden = True    ' ��� �������
            Else
'-- ������� ��� �� ���������� �������� � ����������� �� �����
                Rub = .Cells(i, PAYRUB_COL)
                If Rub >= 1000000 Then
                    .Cells(i, PAYRUB_COL).Interior.Color = rgbBrown
                ElseIf Rub > 500000 Then
                    .Cells(i, PAYRUB_COL).Interior.Color = rgbOrange
                ElseIf Rub > 300000 Then
                    .Cells(i, PAYRUB_COL).Interior.Color = rgbBisque
                ElseIf Rub > 30000 Then
                    .Cells(i, PAYRUB_COL).Interior.Color = rgbBeige
                End If
            End If
            
'-- ������� ���������
            If .Cells(i, PAYDOGOVOR_COL) <> "" Then     ' ������� ��������
                .Cells(i, PAYDOGOVOR_COL).Interior.Color = rgbLightBlue
            End If
            If .Cells(i, PAYOSNDOGOVOR_COL) <> "" Then  ' ������� ���.��������
                .Cells(i, PAYOSNDOGOVOR_COL).Interior.Color = rgbLightBlue
            End If
'-- ������� ������� ADSK � ������������ � SF_PA
            If InStr(.Cells(i, PAYGOOD_COL), "Auto") <> 0 Then  ' Autodesk
                If .Cells(i, PAYADSK_COL) = "" Then
                    .Cells(i, PAYGOOD_COL).Interior.Color = rgbADSK
                Else
                    .Cells(i, PAYGOOD_COL).Interior.Color = rgbPink
                End If
            End If
'-- ������� ������� � - ����������� ���� � SF
''            If .Cells(i, PAYISACC_COL) = "1" Then
''                .Cells(i, PAYISACC_COL).Interior.Color = rgbYellow
''            Else
''                .Cells(i, PAYISACC_COL).Interior.Color = rgbRed
''            End If
            
'-- �������� ���
            Doc = Trim(.Cells(i, PAYDOC_COL))
            If Doc = "" Or InStr(Doc, "��� ���") <> 0 Then .Rows(i).Hidden = True
            
        Next i
'-- �������� ����� � �������1�
        Range("Payment_Summary").Copy Destination:=.Cells(RepTOC.EOL + 1, 1)
    End With
'    ModEnd REP_1C_P_PAINT
End Sub
