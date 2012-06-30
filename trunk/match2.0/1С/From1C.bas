Attribute VB_Name = "From1C"
'---------------------------------------------------------------------------
' ������� ��� �������� ������� �� 1�
'
' From1Cpayment  - �������� ���� ������ �� 1� "������ ����� �� �����"
' From1Cdogovor  - �������� ���� ������ 1� "��������" ����� �� 1�
' From1Caccount  - �������� ���� ������ 1� "������� .." ����� �� 1�
' FromStock      - ������ ��������� ����� � ����� Stock
'
' 29.6.2012 �.�.������� match 2.0

Option Explicit
Sub From1Cpayment()
'
' From1Cpayment() - �������� ���� ������ 1� "������ ����� �� �����" ����� ������ �� ������ �����
'
'  25.6.12 - match 2.0
'  27.6.12 - ����������� �� ������ �� �������� �� ������ ����

    Dim LO, Ln, Lines, LinesOld, PaymentsheetName, Pold

    ModStart REP_1C_P_LOAD
    
    Sheets(1).Select                   ' ����� ����� � ����� 1
    Sheets(1).UsedRange.Activate
    With Selection.Font    ' ������������� �� ������ ������ ����������� �����
        .Name = "Calibri"
        .size = 8
    End With
    
    Sheets(2).Columns("A:E").Copy   ' �� �������� ������ �������� ������� A:E
    Sheets(1).Columns("A:A").Select '    � ��������� �� ����� � ������ ������
    Selection.Insert Shift:=xlToRight
        
    Range("J:Q,T:U,W:X").Select           ' ������ ���������� �������� �������
    Selection.EntireColumn.Hidden = True  '   ..�������� ��������, ���������
                                          '   .. ��������,������� � ����
    LinesOld = EOL(2)               ' ���-�� ����� � ������ ������
    Lines = EOL(1)                  ' ���-�� ����� � ����� ������

' ��������� ������� ������ �� ����� ������� �������
    LO = LinesOld - 3
    Ln = Lines - 2
    Range(Cells(LO, 1), Cells(LO, 5)).Select
    If Ln > LO Then
        Selection.AutoFill Destination:=Range(Cells(LO, 1), Cells(Ln, 5)), _
            Type:=xlFillDefault
    End If
    
    Call DateCol(PAY_SHEET, PAYDATE_COL) ' �������������� ������� ���
    SheetSort PAY_SHEET, PAYDATE_COL     ' ��������� ������� �� ���� ������� �����
        
    Sheets(2).Select                ' ��������� ������ ����� - ����� ������
    Range(Cells(LinesOld - 2, 2), Cells(LinesOld, 18)).Copy
    Sheets(1).Select
    Range(Cells(Lines - 1, 2), Cells(Lines - 1, 2)).Activate
    ActiveSheet.Paste
    
    SFmatchFill PAY_SHEET   '*** ��������� Update �� ������� SF
  
'*******************************************************
    Call PaymentPaint    '* ������������ ���� �������� *
'*******************************************************
    
' ����������� �������� ������ ������ � ���� SF
    PaymentsheetName = Sheets(1).Name   ' ����� ������ ������� � ������
    Pold = Sheets(2).Name               '   ..������� �� ��������
    
    SheetsCtrlH "SF", Pold, PaymentsheetName
    SheetsCtrlH "P_PaidContract", Pold, PaymentsheetName
'    SheetsCtrlH "P_Update", Pold, PaymentsheetName
        
    Sheets(2).Select
    ActiveWindow.SelectedSheets.Delete  ' ������� ������ ��������� �����
    
    Sheets(PaymentsheetName).Select
    With ActiveWorkbook.Sheets(1).Tab   ' ��� ������ ������ - �������
        .Color = 255
        .TintAndShade = 0
    End With
    
    Doing = "����� �������� " & Sheets(1).Name
    ModEnd 1
End Sub
Sub SFmatchFill(SheetN)
'
' - SFmatchFill SheetN
'       ���������� ������� PAYINSF_COL � PAYSFLN_COL ����� PAY_SHEET
'       � ����������� �� ������� �������� � ������� SF_COD_COL ����� SF
' 30.6.12

    Dim L As Integer        '= EOL ������������ ������
    Dim iPaid As Integer    '= ����� "������" ������ � SF
    Dim AccCol As Integer   '= ����� "���������" �������
    Dim i As Integer        '= ����� ������� ������ ������������ ������
    Dim j As Integer        '= ����� ������ SF_PA
    
'    SFaccColFill PAY_SHEET  ' � ������� 1 ���� ����������� ���� � SF
'    SFaccCol PAY_SHEET, PAY_RESLINES    ' ������������ ������� A
        
    Select Case SheetN
        Case PAY_SHEET:
            L = EOL_PaySheet
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
        
        
    Set DB_SFDC = Workbooks.Open(F_SFDC, UpdateLinks:=False, ReadOnly:=True)
    With ThisWorkbook.Sheets(PAY_SHEET)
        For i = 2 To L
            Progress i / L
                '-- "�����������" ��� �����������  � "�������" ��� � SFacc --
            .Cells(i, AccCol) = Replace(Compressor(.Cells(i, AccCol)), vbCrLf, "")
            
            If IsMatch(.Cells(i, AccCol), SFACC_ACC1C_COL, , SFacc, DB_SFDC) Then
                .Cells(i, PAYISACC_COL) = "1"
            Else
                .Cells(i, PAYISACC_COL) = ""
            End If
                        
                '-- "�������" � �������� � SF --
            If IsMatch(.Cells(i, PAYCODE_COL), SF_COD_COL, iPaid, SF, DB_SFDC) Then
                .Cells(i, PAYINSF_COL) = "1"
                .Cells(i, PAYSFLN_COL) = iPaid
                SFid = DB_SFDC.Sheets(SF).Cells(iPaid, SF_PAYID_COL)
            Else
                .Cells(i, PAYINSF_COL) = ""
                .Cells(i, PAYSFLN_COL) = ""
                SFid = ""
            End If
            
                '-- "�������" � ����������� ADSK � SF_PA --
    '..... ����� ����� ����� ��������� ��� ��������� ADSK �� ������� � "+"
            If IsMatch(SFid, SFPA_PAYID_COL, , SF_PA, DB_SFDC) Then
                .Cells(i, PAYADSK_COL) = "1"
            Else
                .Cells(i, PAYADSK_COL) = ""
            End If
        Next i
    End With
    DB_SFDC.Close SaveChanges:=False
End Sub
Sub testSFmatchFill()
    EOL_PaySheet = EOL(PAY_SHEET) - PAY_RESLINES
    SFmatchFill PAY_SHEET
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
