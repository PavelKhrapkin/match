Attribute VB_Name = "From1C"
'---------------------------------------------------------------------------
' ������� ��� �������� ������� �� 1�
'
' From1Cpayment  - �������� ���� ������ �� 1� "������ ����� �� �����"
' From1Cdogovor  - �������� ���� ������ 1� "��������" ����� �� 1�
' From1Caccount  - �������� ���� ������ 1� "������� .." ����� �� 1�
' FromStock      - ������ ��������� ����� � ����� Stock
' �.�.������� 30.4.2012
'   5.2.2012 - ����� ��������� ������� � �����
'   7.2.2012 - ������� ����� �����?
'   29.2.2012 - bug fix, �������
'   6.3.12 - Options Explicit, ������������� �������� � ����������
'  22.4.12 - ������� From1Caccount
'  17.5.12 - �������� ��������� �����

Option Explicit
Sub From1Cpayment()
'
' �������� ���� ������ 1� "������ ����� �� �����" ����� ������ �� ������ �����
'   15.1.2012
'   10.2.2012 - �������� ������� �������� � �������� �������
'   29.2.2012 - �������, EOL
'   6.3.12 - Option Explicit
'  30.4.12 - SFaccColFill

    Dim LO, Ln, PaymentsheetName, Pold

    ModStart 2, "������� ����� ����� 1� �� �������� � Match SF-1C"

    CheckSheet 1, 1, 1, Stamp1Cpay1    ' ��������� ������������ ������ ������
    CheckSheet 1, 1, 2, Stamp1Cpay2
    CheckSheet 2, 1, 6, Stamp1Cpay1    ' ������� ����� �� �����?
    CheckSheet 2, 1, 7, Stamp1Cpay2
    
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
    
    Sheets(2).Select                ' ��������� ������ ����� - ����� ������
    Range(Cells(LinesOld - 2, 2), Cells(LinesOld, 18)).Copy
    Sheets(1).Select
    Range(Cells(Lines - 1, 2), Cells(Lines - 1, 2)).Activate
    ActiveSheet.Paste
      
    SFaccColFill PAY_SHEET  ' � ������� 1 ���� ����������� ���� � SF
    SFaccCol PAY_SHEET, PAY_RESLINES    ' ������������ ������� A
  
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
