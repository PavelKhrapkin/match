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
'   11.8.12

    Dim R As TOCmatch
    Dim B As Range
    R = GetRep(ActiveSheet.Name)
    B = Range(Cells(1, 1), Cells(3, BIG))
    B.Copy Destination:=Cells(R.EOL + 2, 1)
    B.Delete
End Sub
Sub AccPaint()
'
' - AccPaid() - ������� ������� � - ����������� ���� � SF
'
    Dim i As Long
    Dim RepTo As TOCmatch
    Dim R As Range
    
    RepTo = GetRep(ActiveSheet.Name)
    With Workbooks(RepTo.RepFile).Sheets(RepTo.SheetN)
        For i = 2 To RepTo.EOL
            Progress i / RepTo.EOL
            R = .Cells(i, PAYISACC_COL)
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
'''''Sub From1Cdogovor()
''''''
'''''' �������� ���� ������ 1� "��������" ����� ������ �� ������ �����
''''''
'''''' 2.1.2012 �.�.�������
'''''' 6.1.2012 ��������� ������� � - <��������>/<�������>
'''''' 12.1.2012 ���������� � �������������� �����������
'''''' 20.1.2012 �� �������� - ������� � �������
'''''' 6.3.12 - ��������� �� ����� � �����
'''''' 30.4.12 - SFaccColFill
'''''' 17.4.12 - bug fix � ����� ������� ������ �������
'''''
'''''    Dim Total, BadContr
'''''
'''''    LinesOld = ModStart(4, "��������� �� 1� ����� �� ���������") - DOGRES
'''''
'''''    CheckSheet 4, 1, 10, Stamp1Cdog1    ' ����� ��������
'''''
'''''    Lines = EOL(1) - 2      ' ���-�� ����� - ��������� � ����� ������
'''''
'''''    Sheets(4).Columns("A:H").Copy   ' � ������� ������ �������� ������� A:H
'''''    Sheets(1).Select
'''''    Sheets(1).Columns("A:A").Select '     .. � ��������� �� ����� ������ ������
'''''    Selection.Insert Shift:=xlToRight
'''''                                    ' ���� � ����� ������ �������� ������, ��
'''''    If Lines > LinesOld Then        '     .. ��������� ������� ������ �� �����
'''''        Range(Cells(LinesOld, 1), Cells(LinesOld, 8)).Select
'''''        Selection.AutoFill Destination:= _
'''''            Range(Cells(LinesOld, 1), Cells(Lines, 8)), Type:=xlFillDefault
'''''    Else
'''''        Range(Cells(Lines + 1, 1), Cells(LinesOld, 8)).Select
'''''        Selection.Delete Shift:=xlUp
'''''    End If
'''''
'''''    Sheets(4).Range("A" & LinesOld + 1 & ":O" & LinesOld + DOGRES).Copy
'''''    Sheets(1).Select                ' ��������� ������ ����� - ����� ������
'''''    Range("A" & Lines + 1).Select
'''''    ActiveSheet.Paste               ' ��������� ������ ����� - ����� ������
'''''
'''''    Sheets(1).Move After:=Sheets("SF") ' ���������� ����� ����� ����� ����� SF
'''''    Sheets(3).UsedRange.Activate
'''''    Rows.RowHeight = 15             ' ������ ����� �� ����� = 15
'''''
'''''    SheetsCtrlH "SFD", Sheets(4).Name, Sheets(3).Name ' ��������� ������� � SFD
'''''' >>>>>>>>>>>>>>>> ����� Ctrl/H ��� ������ ������ Match
'''''
'''''    Total = EOL("SFD")
'''''    BadContr = Cells(Total, 1)
'''''
'''''    MsgBox "� SF " & Total - 7 & " ���������, �� ��� " & BadContr _
'''''                                                & " �� ������������� 1�"
'''''    Sheets(4).Delete                ' ������� ������ ����� �� ���������
'''''
'''''    SFaccColFill DOG_SHEET  ' � ������� 1 ���� ����������� ���� � SF
'''''    SFaccCol DOG_SHEET, DOGRES    ' ������������ ������� A
'''''
''''''************************************************************
'''''    Call ContractPaint       '* ������������ ���� ��������� *
''''''************************************************************
'''''
'''''    ActiveWorkbook.Sheets(3).Tab.Color = LimeG  ' Tab ������ ������ - �������
'''''    Doing = Sheets(3).Name
'''''    Call ModEnd(3)
'''''End Sub
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
            If .Cells(i, PAYISACC_COL) = "1" Then
                .Cells(i, PAYISACC_COL).Interior.Color = rgbYellow
            Else
                .Cells(i, PAYISACC_COL).Interior.Color = rgbRed
            End If
            
'-- �������� ���
            Doc = Trim(.Cells(i, PAYDOC_COL))
            If Doc = "" Or InStr(Doc, "��� ���") <> 0 Then .Rows(i).Hidden = True
            
        Next i
'-- �������� ����� � �������1�
        Range("Payment_Summary").Copy Destination:=.Cells(RepTOC.EOL + 1, 1)
    End With
'    ModEnd REP_1C_P_PAINT
End Sub
''Sub From1Caccount()
'''
''' �������� ���� ������ 1� "������ �������� 1�" ����� ������ �� ������ �����
'''
''' 22.4.2012 �.�.�������
''' 30.4.12 - SFaccColFill
''' 16.6.12 - �������, ����� ������� �������-������ �� ������� ������
''
''    ModStart Acc1C, "���������� ����� ����������� �������� 1�"
''
''    LinesOld = EOL(Acc1C) - ACC1C_RES   ' ���-�� ����� � ������ ������
''    Lines = EOL(1) - ACC1C_RES          ' ���-�� ����� � ����� ������
''
''    CheckSheet 1, 4, 2, ACC1C_STAMP
''    CheckSheet Acc1C, 1, A1C_NAME_COL, ACC1C_STAMP
''
''    ' ������c�� ������� ������ � �����
''    Sheets(1).Select
''    Rows("1:" & ACC1C_RES).Cut
''    Cells(Lines + ACC1C_RES + 2, 1).Select
''    ActiveSheet.Paste
''    Rows("1:" & ACC1C_RES).Delete
''
''    ' ��������� ������� 1
''    Sheets(1).Columns("A:A").Select
''    Selection.Insert Shift:=xlToRight
''
''    Sheets(Acc1C).Name = "RRR"          ' ������� ����� ��������������� � RRR
''    Worksheets(1).Name = Acc1C          ' ����� ����� ��������������� � "������ �������� 1�"
''
''    Sheets(Acc1C).UsedRange.Activate
''    Rows.RowHeight = 15                 ' ������ ����� �� ����� = 15
''
''    Sheets(Acc1C).Move After:=Sheets("RRR") ' ���������� ����� ����� ����� ����� Acc1C
''    Sheets("RRR").Delete
''
''    SFaccColFill Acc1C              ' ��������� � ������� 1 ���� ����������� ���� � SF
''    SFaccCol Acc1C, ACC1C_RES       ' ������������ �������: ������ - ���� � SF, ������� - ���
''
''    Sheets(Acc1C).Tab.Color = rgbRed    ' ���������� Tab ������ ������ � ������� ����
''    ModEnd Acc1C
''End Sub
'''Sub FromStock()
''''
'''' (*) FromStock() - ������ ��������� ����� � ����� Stock
''''       17/5/12
''''   18.5.12 - ������� ������� ContrADSK
'''
'''
'''    Dim LO As Integer, Ln As Integer    ' ���-�� ����� � ������ � ����� �������
'''    Dim Msg As String
'''
'''    LO = ModStart(STOCK_SHEET, "���������� ��������� �����")
'''    Ln = EOL(1)
'''
'''    CheckSheet 1, 3, STOCK_PRODUCT_COL - STOCK_MYCOLS, STOCK_STAMP
'''    CheckSheet STOCK_SHEET, 1, STOCK_PRODUCT_COL, STOCK_STAMP
'''
'''    Sheets(1).Select
'''    Rows("1:2").Select        ' ������� ������ ��� ������
'''    Selection.Delete Shift:=xlUp
'''
'''    Sheets(STOCK_SHEET).Columns("A:F").Copy ' �� �������� ������ �������� ������� A:F
'''    Sheets(1).Columns("A:A").Select         '    � ��������� �� ����� � ������ ������
'''    Selection.Insert Shift:=xlToRight
'''
'''    Sheets(STOCK_SHEET).Name = "Tmp"  ' ��������������� ������� ����� � "Tmp"
'''    Sheets(1).Name = STOCK_SHEET      '   .. � ����� � "�����"
'''
'''    Sheets(STOCK_SHEET).Move Before:=Sheets("Tmp")  ' ���������� ����� �����
'''    Sheets("Tmp").Delete                ' ���������� ������� ����� SF
'''    Sheets(STOCK_SHEET).Tab.Color = rgbBlue ' ���������� Tab ������ ������
'''
'''    Msg = "� ������� ��������� ����� " & LO & " �����, � ����� "
'''    If LO = Ln Then Msg = Msg & "���� "
'''    MsgBox Msg & Ln
'''
'''    ModEnd STOCK_SHEET
'''End Sub
