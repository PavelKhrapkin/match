Attribute VB_Name = "From1C"
'---------------------------------------------------------------------------
' ������� ��� �������� ������� �� 1�
'
' S ContractPaint() - ������������ ���� ���������
' - Paint(iStr,Col,Criteria,Color,[Mode]) - ��������� ����� (iStr,Col) � ���� Color
' - Acc1C_Bottom()  - ������� �����x ���� ����� Acc1� � �����
' S AccPaid()       - ������� ������� � - ����������� ���� � SF
'<*> From1Cpayment  - �������� ���� ������ �� 1� "������ ����� �� �����"
' -  SFmatchFill(SheetN)  - ���������� ������ ����� SheetN �� SFDC
' -  CSmatch(Val,Col,[SheetN],[DB]) - Case Sensitive match ���������� ����� ������
'           � Val � ������� Col ����� SheetN � DB. ���� Val �� ������- ���������� 0.
'<*> From1Cdogovor  - �������� ���� ������ 1� "��������" ����� �� 1�
'<*> From1Caccount  - �������� ���� ������ 1� "������� .." ����� �� 1�
'(*) FromStock      - ������ ��������� ����� � ����� Stock
'
'!�! SFlnkFill(RepFr,ColFr,ColFrId, ColVal, ColTo) - "�������" ����� � DocFr � DocTo,
'               ��������� � ������� ColTo Id �������, ���������� �� �������� � ColFr
'
' 8.11.2012 �.�.������� match 2.0

Option Explicit
Sub ContractPaint()
'
' S ContractPaint() - ������������ ���� ���������
' 10.8.12
'  1.9.12 - StepIn
' 14.9.12 - ��������� � Paint
'  4.11.12 - ���� "�� �������������" - �������

    StepIn
    
    Dim D As TOCmatch
    D = GetRep(DOG_SHEET)
    Dim i As Long
    
    ScreenUpdate False
    
    For i = 2 To D.EOL
        ActiveSheet.Rows(i).Interior.Color = rgbWhite
        Paint i, DOGSFSTAT_COL, "������", rgbLightGreen ' �������� �������� � SF- �������
        Paint i, DOGSFSTAT_COL, "������", rgbOrange     ' �������� �������� - ���������
        Paint i, DOGSFSTAT_COL, "��������", rgbLightBlue ' ��������� - �������
        Paint i, DOGSFSTAT_COL, "�� ���������", Antique ' �� ��������� - Antique
        Paint i, DOGSFSTAT_COL, "��� � SF", rgbWhite    ' ��� � SF - �� ����������
        Paint i, DOGPAID1C_COL, "1", LimeG, 1           ' ���������� - ����� �������
        Paint i, DOGISINV1C_COL, "1", rgbOlive, 1       ' ��������� ���� - ���������
        Paint i, DOG1CSCAN_COL, "1", rgbViolet, 1       ' ������������� - ����������
        Paint i, DOG1CSCAN_COL, "0", rgbRed, 1          ' �� ������������� - �������
    Next i
    
    ScreenUpdate True
    
'-- �������� ����� � �������1�
    DB_MATCH.Sheets(Header).Range("HDR_1C_Contract_Summary").Copy _
            Destination:=ActiveSheet.Cells(D.EOL + 1, 1)
End Sub
Sub Paint(iStr As Long, Col As Long, Criteria As String, Color, Optional Mode As Integer = 0)
'
' - Paint(iStr,Col,Criteria,Color,[Mode]) - ��������� ����� (iStr,Col) � ���� Color
'                            ��� �������� Criteria, ��� ��� ������ ���� ������� Mode=1
' 14.9.12

    Const DOG_COLS = 26     ' ����� ������� � ������� ���������
    
    With ActiveSheet
        If .Cells(iStr, Col) = Criteria Then
            If Mode = 1 Then
                .Cells(iStr, Col).Interior.Color = Color
            Else
                Range(Cells(iStr, 2), Cells(iStr, DOG_COLS)).Interior.Color = Color
            End If
        End If
    End With
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
' S AccPaid() - ������� ������� � - ����������� ���� � SF
'   14.8.12
'   31.8.12 - ��������� StepIn

    StepIn

    Dim i As Long
    Dim RepTo As TOCmatch
    Dim R As Range
    
    ScreenUpdate False
    
    RepTo = GetRep(ActiveSheet.Name)
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
    
    ScreenUpdate True
End Sub

Sub testCSmatch()
    If "G" = "g" Then Stop
    Dim a
    ThisWorkbook.Sheets("Sheet1").Select
    a = CSmatch("g12", 1)
    a = CSmatch("g121", 1)
    
    ModStart REP_1C_P_PAINT
    Set DB_SFDC = Workbooks.Open(F_SFDC, UpdateLinks:=False, ReadOnly:=True)
    DB_SFDC.Sheets(SFacc).Select
    a = CSmatch("��� ""���""", 2)
    ModEnd
End Sub
Sub PaymentPaint(ByVal BottomHDR As String)
'
' - PaymentPaint(BottomHDR) - ������������ ���� �������� 1C � ������ BottomHDR
' 24.6.12 ���������� ��� match 2.0
'  7.8.12 ��������� ��� ���
' 31.8.12 - ��������� StepIn
'  7.2.13 - �������� BottomHDR; ������� ���� ������, ���������� � SF

    StepIn

    Dim i As Integer
    Dim Rub, Doc    '���� "����� ���" � "����.���"
    
    Range("A1:AC" & RepTOC.EOL).Interior.Color = rgbWhite   ' ���������� �������
    Rows("2:" & RepTOC.EOL).RowHeight = 15    ' ������ ����� �� ����� = 15
    
    With Sheets(PAY_SHEET)
        For i = 2 To RepTOC.EOL
            Progress i / RepTOC.EOL
            If .Cells(i, PAYINSF_COL) = 1 Then          ' ������� ������� � SF
'                Range(Cells(i, 2), Cells(i, AllCol)).Interior.Color = rgbLightGreen
                Range(Cells(i, 2), Cells(i, .Columns.Count)).Interior.Color = rgbLightGreen
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
            
'-- �������� ���
            Doc = Trim(.Cells(i, PAYDOC_COL))
            If Doc = "" Or InStr(Doc, "��� ���") <> 0 Then .Rows(i).Hidden = True
        Next i

'-- �������� ����� � �������1�
        DB_MATCH.Sheets(Header).Range(BottomHDR).Copy Destination:=.Cells(RepTOC.EOL + 1, 1)
    End With
    
End Sub
Sub SFlnkFill(DocFr, ColFr, ColFrId, ColVal, ColTo)
'
' � SFlnkFill(RepFr,ColFr,ColFrId, ColVal, ColTo) - "�������" ����� � DocFr � DocTo,
'       ��������� � ������� ColTo Id �������, ���������� �� �������� � ColFr
' 8.8.12
' 31.8.12 - ��������� StepIn
'  3.11.12 -- ������� - �������� ���������

    StepIn
    
    Dim DocTo As String ' ��� �������� ��������� - ������
    Dim RepFr As TOCmatch, RepTo As TOCmatch
    Dim Val
    Dim i As Integer, N As Integer
    
    DocTo = ActiveSheet.Name
    RepTo = GetRep(ActiveSheet.Name)
    Workbooks(RepTo.RepFile).Sheets(RepTo.SheetN).Activate
    RepFr = GetRep(DocFr)
    With Workbooks(RepTo.RepFile).Sheets(RepTo.SheetN)
        Workbooks(RepFr.RepFile).Sheets(RepFr.SheetN).Activate
        For i = 2 To RepTo.EOL
            Progress i / RepTo.EOL                                  'ActiveSheet = SFacc
            Val = .Cells(i, ColVal)                                 '������ ���� � 1C -- Val
            N = CSmatch(Val, ColFr)
            If N > 0 Then
'                .Cells(i, ColTo) = Workbooks(RepFr.RepFile).Sheets(RepFr.SheetN).Cells(N, ColFrId)
                .Cells(i, ColTo) = ActiveSheet.Cells(N, ColFrId)
            Else
                .Cells(i, ColTo) = ""
            End If
        Next i
    End With
End Sub

