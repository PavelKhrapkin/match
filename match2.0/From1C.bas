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
' 25.7.2012 �.�.������� match 2.0

Option Explicit
Sub From1Cpayment()
'
' From1Cpayment() - �������� ���� ������ 1� "������ ����� �� �����" ����� ������ �� ������ �����
'
'  25.6.12 - match 2.0
'  27.6.12 - ����������� �� ������ �� �������� �� ������ ����
'  25.7.12 - InsMyCol � InsSummary

    ModStart REP_1C_P_LOAD
    
    InsMyCol "Payment_MyCol"
    InsSummary "Payment_Summary"

    Range("J:Q,T:U,W:X").Select           ' ������ ���������� �������� �������
    Selection.EntireColumn.Hidden = True  '   ..�������� ��������, ���������
                                          '   .. ��������,������� � ����
    Call DateCol(PAY_SHEET, PAYDATE_COL) ' �������������� ������� ���
'    SheetSort PAY_SHEET, PAYDATE_COL     ' ��������� ������� �� ���� ������� �����
    
    SFmatchFill PAY_SHEET   '*** ��������� Update �� ������� SF

'*******************************************************
    Call PaymentPaint    '* ������������ ���� �������� *
'*******************************************************

    ModEnd
End Sub
Sub testSFmatchFill()
    ModStart REP_1C_P_PAINT
    Set DB_SFDC = Workbooks.Open(F_SFDC, UpdateLinks:=False, ReadOnly:=True)
    SFmatchFill PAY_SHEET
    ModEnd
End Sub
Sub SFmatchFill(SheetN)
'
' - SFmatchFill(SheetN)  - ���������� ������ ����� SheetN �� SFDC
' 8.7.12

    Dim L As Integer        '= EOL ������������ ������
    Dim Acc As String       '���� ������ 1� � �������������
    Dim SFid As String      '���� IdSF - Id ������� � SF
    Dim iPaid As Integer    '= ����� "������" ������ � SF
    Dim AccCol As Integer   '= ����� "���������" �������
    Dim i As Integer        '= ����� ������� ������ ������������ ������
    
'    SFaccColFill PAY_SHEET  ' � ������� 1 ���� ����������� ���� � SF
'    SFaccCol PAY_SHEET, PAY_RESLINES    ' ������������ ������� A
        
    Select Case SheetN
        Case PAY_SHEET:
'            L = EOL_PaySheet
            L = RepTOC.EOL
            AccCol = PAYACC_COL
        Case DOG_SHEET:
            L = EOL_DogSheet
            AccCol = DOG1CACC_COL
        Case Acc1C:
            L = EOL_Acc1C
            AccCol = A1C_NAME_COL
        Case Else
            ErrMsg FATAL_ERR, "������������ ���� � SFaccColFill"
            Stop
    End Select
        
    DB_SFDC.Sheets(SFacc).Activate
'    DB_SFDC.Sheets(SFacc).Select
    With ThisWorkbook.Sheets(PAY_SHEET)
        For i = 2 To L
            Progress i / L / 3
                '-- "�����������" ��� �����������  � "�������" ��� � SFacc --
            Acc = Replace(Compressor(.Cells(i, AccCol)), vbCrLf, "")
            .Cells(i, AccCol) = Acc
            If CSmatch(Acc, SFACC_ACC1C_COL) <> 0 Then
                .Cells(i, PAYISACC_COL) = "1"
            Else
                .Cells(i, PAYISACC_COL) = ""
            End If
        Next i
                        
        DB_SFDC.Sheets(SF).Select
        For i = 2 To L
            Progress 1 / 3 + i / L / 3
                '-- "�������" � �������� � SF --
            iPaid = CSmatch(.Cells(i, PAYCODE_COL), SF_COD_COL)
            If iPaid <> 0 Then
                .Cells(i, PAYINSF_COL) = "1"
                SFid = DB_SFDC.Sheets(SF).Cells(iPaid, SF_PAYID_COL)
                .Cells(i, PAYIDSF_COL) = SFid
            Else
                .Cells(i, PAYINSF_COL) = ""
                .Cells(i, PAYIDSF_COL) = ""
            End If
        Next i

        DB_SFDC.Sheets(SF_PA).Select
        For i = 2 To L
            Progress 2 / 3 + i / L / 3
                '-- "�������" � ����������� ADSK � SF_PA --
    '..... ����� ����� ����� ��������� ��� ��������� ADSK �� ������� � "+"
            SFid = .Cells(i, PAYIDSF_COL)
            If CSmatch(SFid, SFPA_PAYID_COL) <> 0 Then
                .Cells(i, PAYADSK_COL) = "1"
            Else
                .Cells(i, PAYADSK_COL) = ""
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
Sub From1Cdogovor()
'
' �������� ���� ������ 1� "��������" ����� ������ �� ������ �����
'
' 2.1.2012 �.�.�������
' 6.1.2012 ��������� ������� � - <��������>/<�������>
' 12.1.2012 ���������� � �������������� �����������
' 20.1.2012 �� �������� - ������� � �������
' 6.3.12 - ��������� �� ����� � �����
' 30.4.12 - SFaccColFill
' 17.4.12 - bug fix � ����� ������� ������ �������
    
    Dim Total, BadContr

    LinesOld = ModStart(4, "��������� �� 1� ����� �� ���������") - DOGRES
    
    CheckSheet 4, 1, 10, Stamp1Cdog1    ' ����� ��������
    
    Lines = EOL(1) - 2      ' ���-�� ����� - ��������� � ����� ������
    
    Sheets(4).Columns("A:H").Copy   ' � ������� ������ �������� ������� A:H
    Sheets(1).Select
    Sheets(1).Columns("A:A").Select '     .. � ��������� �� ����� ������ ������
    Selection.Insert Shift:=xlToRight
                                    ' ���� � ����� ������ �������� ������, ��
    If Lines > LinesOld Then        '     .. ��������� ������� ������ �� �����
        Range(Cells(LinesOld, 1), Cells(LinesOld, 8)).Select
        Selection.AutoFill Destination:= _
            Range(Cells(LinesOld, 1), Cells(Lines, 8)), Type:=xlFillDefault
    Else
        Range(Cells(Lines + 1, 1), Cells(LinesOld, 8)).Select
        Selection.Delete Shift:=xlUp
    End If
    
    Sheets(4).Range("A" & LinesOld + 1 & ":O" & LinesOld + DOGRES).Copy
    Sheets(1).Select                ' ��������� ������ ����� - ����� ������
    Range("A" & Lines + 1).Select
    ActiveSheet.Paste               ' ��������� ������ ����� - ����� ������
    
    Sheets(1).Move After:=Sheets("SF") ' ���������� ����� ����� ����� ����� SF
    Sheets(3).UsedRange.Activate
    Rows.RowHeight = 15             ' ������ ����� �� ����� = 15
    
    SheetsCtrlH "SFD", Sheets(4).Name, Sheets(3).Name ' ��������� ������� � SFD
' >>>>>>>>>>>>>>>> ����� Ctrl/H ��� ������ ������ Match

    Total = EOL("SFD")
    BadContr = Cells(Total, 1)
    
    MsgBox "� SF " & Total - 7 & " ���������, �� ��� " & BadContr _
                                                & " �� ������������� 1�"
    Sheets(4).Delete                ' ������� ������ ����� �� ���������
    
    SFaccColFill DOG_SHEET  ' � ������� 1 ���� ����������� ���� � SF
    SFaccCol DOG_SHEET, DOGRES    ' ������������ ������� A
      
'************************************************************
    Call ContractPaint       '* ������������ ���� ��������� *
'************************************************************

    ActiveWorkbook.Sheets(3).Tab.Color = LimeG  ' Tab ������ ������ - �������
    Doing = Sheets(3).Name
    Call ModEnd(3)
End Sub
Sub From1Caccount()
'
' �������� ���� ������ 1� "������ �������� 1�" ����� ������ �� ������ �����
'
' 22.4.2012 �.�.�������
' 30.4.12 - SFaccColFill
' 16.6.12 - �������, ����� ������� �������-������ �� ������� ������

    ModStart Acc1C, "���������� ����� ����������� �������� 1�"

    LinesOld = EOL(Acc1C) - ACC1C_RES   ' ���-�� ����� � ������ ������
    Lines = EOL(1) - ACC1C_RES          ' ���-�� ����� � ����� ������
    
    CheckSheet 1, 4, 2, ACC1C_STAMP
    CheckSheet Acc1C, 1, A1C_NAME_COL, ACC1C_STAMP
    
    ' ������c�� ������� ������ � �����
    Sheets(1).Select
    Rows("1:" & ACC1C_RES).Cut
    Cells(Lines + ACC1C_RES + 2, 1).Select
    ActiveSheet.Paste
    Rows("1:" & ACC1C_RES).Delete
    
    ' ��������� ������� 1
    Sheets(1).Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
        
    Sheets(Acc1C).Name = "RRR"          ' ������� ����� ��������������� � RRR
    Worksheets(1).Name = Acc1C          ' ����� ����� ��������������� � "������ �������� 1�"
    
    Sheets(Acc1C).UsedRange.Activate
    Rows.RowHeight = 15                 ' ������ ����� �� ����� = 15
    
    Sheets(Acc1C).Move After:=Sheets("RRR") ' ���������� ����� ����� ����� ����� Acc1C
    Sheets("RRR").Delete
    
    SFaccColFill Acc1C              ' ��������� � ������� 1 ���� ����������� ���� � SF
    SFaccCol Acc1C, ACC1C_RES       ' ������������ �������: ������ - ���� � SF, ������� - ���
      
    Sheets(Acc1C).Tab.Color = rgbRed    ' ���������� Tab ������ ������ � ������� ����
    ModEnd Acc1C
End Sub
Sub FromStock()
'
' (*) FromStock() - ������ ��������� ����� � ����� Stock
'       17/5/12
'   18.5.12 - ������� ������� ContrADSK


    Dim LO As Integer, Ln As Integer    ' ���-�� ����� � ������ � ����� �������
    Dim Msg As String
    
    LO = ModStart(STOCK_SHEET, "���������� ��������� �����")
    Ln = EOL(1)

    CheckSheet 1, 3, STOCK_PRODUCT_COL - STOCK_MYCOLS, STOCK_STAMP
    CheckSheet STOCK_SHEET, 1, STOCK_PRODUCT_COL, STOCK_STAMP
    
    Sheets(1).Select
    Rows("1:2").Select        ' ������� ������ ��� ������
    Selection.Delete Shift:=xlUp
    
    Sheets(STOCK_SHEET).Columns("A:F").Copy ' �� �������� ������ �������� ������� A:F
    Sheets(1).Columns("A:A").Select         '    � ��������� �� ����� � ������ ������
    Selection.Insert Shift:=xlToRight
                                                                
    Sheets(STOCK_SHEET).Name = "Tmp"  ' ��������������� ������� ����� � "Tmp"
    Sheets(1).Name = STOCK_SHEET      '   .. � ����� � "�����"

    Sheets(STOCK_SHEET).Move Before:=Sheets("Tmp")  ' ���������� ����� �����
    Sheets("Tmp").Delete                ' ���������� ������� ����� SF
    Sheets(STOCK_SHEET).Tab.Color = rgbBlue ' ���������� Tab ������ ������
    
    Msg = "� ������� ��������� ����� " & LO & " �����, � ����� "
    If LO = Ln Then Msg = Msg & "���� "
    MsgBox Msg & Ln
        
    ModEnd STOCK_SHEET
End Sub
