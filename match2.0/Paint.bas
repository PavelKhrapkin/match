Attribute VB_Name = "Paint"
'------------------------------------------------------------------------------------
' �������������� � ������������� ������ ��������, ��������� � ������� 1 � 1C.xlsm
'(*) PaymentPaint   - ������������ � ����������� ���� �������� �� 1�
'(*) ContractPaint  - ������������ � ����������� ���� ���������
' T TestAccFill()   - ������� SFaccColFill ��� ����� ��������
' - SFaccColFill(SheetN)    - ��������� ������� 1 ����� SheetN "1",
'                             ���� ����������� � AccCol ���� � SF
' -  SFaccCol(SheetN, ResLines)  - ��������� ������� 1 �� SFacc
' -  ADSK_P_Paint() - ��������� �������� ADSK, � ������ ������� ����� SF_PA
' -  IsP_AbyN(Nstr) - ���������� TRUE, ���� ������ Nstr ������� �������� � ADSK
'
' 30.6.2012 match2.0

Option Explicit

Sub PaymentPaint()
Attribute PaymentPaint.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' - PaymentPaint() - ������������ ���� �������� 1C
' 24.6.12 ���������� ��� match 2.0

    Dim i As Integer
    Dim Rub             '���� "����� ���"

    ModStart REP_1C_P_PAINT    ' ����� "��������� �������� 1�"
     
    Worksheets(PAY_SHEET).Columns("A:AC").Select
    Selection.EntireColumn.Hidden = False ' ��������� ��� ��������� (Hide) ������
    Range("J:Q,T:U,W:X").Select          ' ������ ���������� �������� �������
    Selection.EntireColumn.Hidden = True '�������� ��������, ��������� ��������,
                                         '������� � ����
' ���������� ������ ������� ��� �������� ���������� ������ � SF
    Columns("G:G").ColumnWidth = 8      ' ���� ������� �����
    Columns("F:F").ColumnWidth = 13     ' ��������� ��������
    Columns("H:H").ColumnWidth = 7.29   ' ����
    Columns("Y:Y").ColumnWidth = 8.5    ' �������
    Columns("Z:Z").ColumnWidth = 14     ' �������� �������
    
    Range("A1:AC" & EOL_PaySheet).Interior.Color = rgbWhite   ' ���������� �������
    Rows("2:" & EOL_PaySheet).RowHeight = 15    ' ������ ����� �� ����� = 15
    
    With Sheets(PAY_SHEET)
        For i = 2 To EOL_PaySheet
            Progress i / EOL_PaySheet
            If .Cells(i, PAYINSF_COL) = 1 Then          ' ������� ������� � SF
                Range(Cells(i, 2), Cells(i, AllCol)).Interior.Color = rgbLightGreen
            ElseIf Trim(.Cells(i, PAYDOC_COL)) = "" Or Trim(.Cells(i, PAYSALE_COL)) = "" Then
                .Cells(i, 1).EntireRow.Hidden = True    ' ��� �������
            Else
'-- ������� ��� �� ���������� �������� � ����������� �� �����
                Rub = .Cells(i, PAYRUB_COL)
                If Rub >= 1000000 Then
                    PntCell .Cells(i, PAYRUB_COL), rgbBrown
                ElseIf Rub > 500000 Then
                    PntCell .Cells(i, PAYRUB_COL), rgbOrange
                ElseIf Rub > 300000 Then
                    PntCell .Cells(i, PAYRUB_COL), rgbBisque
                ElseIf Rub > 30000 Then
                    PntCell .Cells(i, PAYRUB_COL), rgbBeige
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
            
        Next i
    End With
    ModEnd REP_1C_P_PAINT
End Sub
Sub PntCell(CelRef, Color)
'
' - PntCell(CelRef, Color) - ������� ������
'   24.6.12

    CelRef.Interior.Color = Color
End Sub
Sub ContractPaint()
Attribute ContractPaint.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' ������������ ���� ���������
'   26.1.2012
'   1.3.12 -  ��������� ������� ���������
'  30.4.12 - SFaccColFill
'   1.5.12 - ������ ContrToDl � ������ ��������� ����� ��������� � SF

    Const reslines = 7  ' ���������� ����� � ������ ����� ������
    
    Lines = ModStart(3, "������������ ���� ���������")
    
'*** ������� ������ ��������������� ������� ������ ����� ���������
    SheetSort 3, 4          ' ��������� ���� �������� �� ������� ������� � ������� �
    
    ' ��� ������� ����� ��������� - �������
    ActiveSheet.UsedRange.EntireColumn.Hidden = False

    Rows("1:" & Lines).Font.Name = "Calibri"    ' ���� ����� � ������� Calibri 8
    Rows("1:" & Lines).Font.size = 8
    
    Range("T" & Lines - 5 & ":V" & Lines - 5).NumberFormat = "#,##0"
    
' ���������� ������ ������� ��� �������� ���������� ������ � SF
    Columns("A:A").ColumnWidth = 1.86   ' SF ��� - ��� ����������� � SF
    Columns("B:B").ColumnWidth = 7.86   ' ������ � SF
    Columns("C:C").ColumnWidth = 6.57   ' ���� ����������� � SF <!> ���� �� �� �������� 2.1.12
    Columns("D:D").ColumnWidth = 6.57   ' ����� � SF
    Columns("E:E").ColumnWidth = 1.71   ' �������� - �������� �� ����� � 1�
    Columns("F:F").ColumnWidth = 1.57   ' ���������� �����
    Columns("G:G").ColumnWidth = 1.43   ' �������������
    Columns("H:H").ColumnWidth = 0.75   ' <��������>/<�������>
    Columns("I:I").ColumnWidth = 0.75
    Columns("J:J").ColumnWidth = 8.45   ' ����� �������� � 1�
    Columns("K:K").ColumnWidth = 0.75
    Columns("L:L").ColumnWidth = 9      ' ���� ����������
    Columns("M:M").ColumnWidth = 7.57   ' ���� �����
    Columns("N:N").ColumnWidth = 10.29  ' ��������
    Columns("O:O").ColumnWidth = 20.57  ' �������� � 1�
    Columns("P:P").ColumnWidth = 12.75  ' ���������
    Columns("Q:Q").ColumnWidth = 15.86  ' ��� ��������
    Columns("R:R").ColumnWidth = 8.14   ' �������� �������
    Columns("S:S").ColumnWidth = 5.57   ' ����� � ���������
    Columns("T:T").ColumnWidth = 14     ' ����� (����)
    Columns("U:U").ColumnWidth = 10     ' ����� �� ������
    Columns("V:V").ColumnWidth = 10     ' ����� �� ��������
    Columns("W:W").ColumnWidth = 5.14   ' ������
    Columns("X:X").ColumnWidth = 11.43  ' ����� � �������
    Columns("Y:Y").ColumnWidth = 24.57  ' ����������
    
    Lines = Lines - reslines    ' � ���������� � �������� �� ��������
    
    Pnt 2, "������", rgbLightGreen      ' �������� �������� � SF - �������
    Pnt 2, "������", rgbOrange          ' �������� �������� - ���������
    Pnt 2, "��������", rgbLightBlue     ' ��������� - �������
    Pnt 2, "�� ���������", Antique      ' �� ��������� - Antique
    Pnt 5, 1, LimeG, 1                  ' ���������� - ����� �������
    Pnt 6, 1, rgbOlive, 1               ' ��������� ���� - ���������
    Pnt 7, 1, rgbViolet, 1              ' ��������� ���� - ���������
    
    ModEnd 3
End Sub
Sub TestAccFill()
'
' T TestAccFill() - ������� SFaccColFill ��� ����� ��������
'   26.6.2012

    SFaccColFill PAY_SHEET  ' � ������� 1 ���� ����������� ���� � SF
    SFaccCol PAY_SHEET, PAY_RESLINES    ' ������������ ������� A
End Sub
Sub SFaccColFill(SheetN)
'
' - SFaccColFill(SheetN)    - ��������� ������� 1 ����� SheetN "1",
'                             ���� ����������� � AccCol ���� � SF
' ��� Case Sensitive ��� ���� �����������, � ������� �� VLOOKUP
'   30.4.12
'  1.5.12 - �������� �� ������ ���� ������ � ��������
'  3.5.12 - ������������� Hash ������ �.����
' 25.6.12 - match 2.0
'  2.7.12 - CSmatch(Acc, SFACC_ACC1C_COL, SFacc, DB_SFDC)

'    Dim SFDC As workbook
     
    Dim L, i, j, AccCol As Integer
    Dim Acc As String
    Dim accInSF As String
    Dim t0, t1
    
''''''''''    Static accHTable(0 To 5000) As String
''''''''''    Static accVTable(0 To 5000) As String
''''''''''
''''''''''    hashInit accHTable, accVTable
    
    t0 = Timer
       
    ModStart REP_1C_SFACCFIL

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

 ' -- ������ �� ����� ����������� � SF SFacc ��� ������ ������ SheetN
'    Set DB_SFDC = Workbooks.Open(F_SFDC, UpdateLinks:=True, ReadOnly:=True)
    With ThisWorkbook.Sheets(SheetN)
        For i = 2 To L
            Progress i / L

            Acc = Replace(Compressor(.Cells(i, AccCol)), vbCrLf, "")
            If Acc = "" Then
                accInSF = ""
            Else
                accInSF = hashGet(accHTable, accVTable, Acc)
            End If
            If accInSF = "$" Then
                .Cells(i, 1) = ""
                If Acc <> "" Then
                    For j = 2 To EOL_SFacc
'                        If Acc = Replace(DB_SFDC.Sheets(SFacc).Cells(j, SFACC_ACC1C_COL), "" & vbCrLf, "") Then
                        If CSmatch(Acc, SFACC_ACC1C_COL, SFacc, DB_SFDC) Then
                            .Cells(i, 1) = "1"
                            Exit For
                        End If
                    Next j
                End If
                hashSet accHTable, accVTable, Acc, .Cells(i, 1)
            Else
                .Cells(i, 1) = accInSF
            End If
        Next i
    End With
    DB_SFDC.Close SaveChanges:=False


    t1 = Timer - t0
    LogWr "[SFaccColFill] Time =" & t1

    ModEnd
End Sub
Sub SFaccCol(SheetN, Optional reslines As Integer = 0)
'
' ������������ ������� �:� � ����� SheetsN �� SFacc - ����������� �����������
' ���� ������ ����������� ���� � SF (�������� � ������� =1)��� ������, ����� �������
'   12.1.2012
'   25.1.2012 - ���������� ������ � Lines; ������ ������������ Private L
'   29.2.2012 - bug fix, EOL ������ AutoFulterReset
'   13.5.2012 - ����� ������� ��������� �����

    Dim L
    
    L = EOL(SheetN) - reslines
    
    Range("B2:AZ" & L).Interior.Color = rgbWhite    ' ���������� ������ ��������� �����
    Range("A2:A" & L).Interior.Color = rgbYellow
    Range("A1:A" & L).AutoFilter Field:=1, Criteria1:="="
    Range("A2:A" & L).Interior.Color = rgbRed  ' ����������� SF ����������� � ������� 1 �������
    ActiveSheet.UsedRange.AutoFilter Field:=1
    
    Rows("1:1").RowHeight = 45          ' ������ ��� ������� ������
    Rows("2:" & L).RowHeight = 15       ' ������ ����� �� ����� = 15
    
End Sub
Sub ADSK_P_Paint()
'
' - ADSK_P_Paint()  - ��������� �������� ADSK, � ������ ������� ����� SF_PA
'   18.5.12

    Dim i As Long
    Dim j As Long
    
    With Sheets(PAY_SHEET)
        For i = 2 To EOL_PaySheet
            If .Cells(i, PAYINSF_COL) = 1 Then
                If GoodType(.Cells(i, PAYGOOD_COL)) = WE_GOODS_ADSK Then
                    If IsP_AbyN(i - 1) Then
                        .Cells(i, PAYGOOD_COL).Interior.Color = rgbPink
                    End If
                End If
            End If
        Next i
    End With
    
End Sub

Function IsP_AbyN(Nstr) As Boolean
'
' - IsP_AbyN(Nstr)  - ���������� TRUE, ���� ������ Nstr ������� �������� � ADSK
'   18.5.2012

    Dim SFid As String
    Dim ADSKid As String
    Dim i As Long

    IsP_AbyN = False
    SFid = "": ADSKid = ""
    On Error Resume Next
    SFid = WorksheetFunction.VLookup(Nstr, _
        Sheets(SF).Range("A:R"), 18, False)
    On Error GoTo 0
    If SFid = "" Then Exit Function
    For i = 2 To EOL_SFlnkADSK
        If Sheets(SF_PA).Cells(i, 2) = SFid Then
            IsP_AbyN = True
            Exit Function
        End If
    Next i
End Function
