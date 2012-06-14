Attribute VB_Name = "Paint"
'-------------------------------------------------------------------------------
' �������������� � ������������� ������ �������� � ��������� � MatchSF-1C.xlsm
'(*) PaymentPaint   - ������������ � ����������� ���� �������� �� 1�
'(*) ContractPaint  - ������������ � ����������� ���� ���������
' -  SFaccCol(SheetN, ResLines)  - ��������� ������� 1 �� SFacc
'
' 27.5.2012
Option Explicit

Sub PaymentPaint()
Attribute PaymentPaint.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' ������������ ���� ��������
'
' * ������ <����> �� ������ ��������
'
'   20.4.2012
'   30.4.12 - SFaccColFill

    Const Reslines = 3  ' ����� ����� � ������ ��� �������

    Lines = ModStart(1, "������������ ���� ��������") - Reslines
     
'*** ������� ������ ��������������� ������� ������ ����� ��������
    SheetSort 1, 2          ' ��������� ���� �������� �� ������� ������� � ������� �
    
    Worksheets(1).Columns("A:AC").Select
    Selection.EntireColumn.Hidden = False ' ��������� ��� ��������� (Hide) ������
    
    Range("J:Q,T:U,W:X").Select          ' ������ ���������� �������� �������
    Selection.EntireColumn.Hidden = True '�������� ��������, ��������� ��������,
                                         '������� � ����
' ���������� ������ ������� ��� �������� ���������� ������ � SF
    Columns("G:G").ColumnWidth = 6.29   ' ���� ������� �����
    Columns("F:F").ColumnWidth = 11.57  ' ��������� ��������
    Columns("H:H").ColumnWidth = 7.29   ' ����
    Columns("Y:Y").ColumnWidth = 8.5    ' �������
    Columns("Z:Z").ColumnWidth = 14     ' �������� �������
    
' ��������� ���� ������� ����� �� ���������� �������
    Range("G2:G" & Lines).Select
    
    Selection.NumberFormat = "dd/mm/yy;@"
'    Selection.NumberFormat = "@"
    Columns("G:G").Select
    Selection.Replace What:=".", Replacement:=".", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

' �������� � ������ ������������� � SF ������� � ���������� ��, ��� ���� ��������
    ActiveSheet.Range("A1:R" & Lines).AutoFilter Field:=4, Criteria1:="="
    
    With Range("R1:R" & Lines)
        .AutoFilter Field:=18, _
            Criteria1:=">30000", Operator:=xlAnd, Criteria2:="<=300000"
            With .Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent6
                .PatternTintAndShade = 0
                .pattern = xlSolid
            End With
        .Interior.TintAndShade = 0.599993896298105  ' �� 30k �� 300k- �������
        .AutoFilter Field:=18
            
        .AutoFilter Field:=18, _
            Criteria1:=">300000", Operator:=xlAnd, Criteria2:="<=500000"
        .Interior.TintAndShade = 0.399975585192419  ' �� 300k �� 500k- �������
        .AutoFilter Field:=18

        .AutoFilter Field:=18, _
            Criteria1:=">500000", Operator:=xlAnd, Criteria2:="<1000000"
        .Interior.TintAndShade = -0.249977111117893 ' �� 0,5 �� 1 ��� - ���������
        .AutoFilter Field:=18

        .AutoFilter Field:=18, Criteria1:=">=1000000"   ' ������� ����� 1 ���.�.
        .Interior.Color = 13209                         '     ���������� ����������
    End With
    ActiveSheet.AutoFilterMode = False
    Rows("1:1").AutoFilter

    Pnt 4, 1, rgbLightGreen       ' ������� ���������� � SF - �������
    Pnt 25, "<>", rgbLightBlue, 1 ' �������� � ������� Y �������
    Pnt 26, "<>", rgbLightBlue, 1 ' ���.�������� � ������� Z �������
    Pnt 19, "=*Auto*", rgbADSK, 1 ' ������ Autodesk
    
    Call ADSK_P_Paint             ' ��������� �������� � ������ SF_PA

    Call DateCol(1, 7)            ' �������������� ������� 7 � ����
    
'    Rows("1:1").Interior.Color = RGB(255, 255, 255) '������ ������ ��� �������
    
    Application.StatusBar = "������� " & Sheets(1).Name & " ������ � ������!"

End Sub
Sub ContractPaint()
Attribute ContractPaint.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' ������������ ���� ���������
'   26.1.2012
'   1.3.12 -  ��������� ������� ���������
'  30.4.12 - SFaccColFill
'   1.5.12 - ������ ContrToDl � ������ ��������� ����� ��������� � SF

    Const Reslines = 7  ' ���������� ����� � ������ ����� ������
    
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
    
    Lines = Lines - Reslines    ' � ���������� � �������� �� ��������
    
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
    SFaccColFill PAY_SHEET  ' � ������� 1 ���� ����������� ���� � SF
    SFaccCol PAY_SHEET, PAY_RESLINES    ' ������������ ������� A
End Sub
Sub SFaccColFill(SheetN)
'
' ��������� ������� 1 ����� SheetN 1 ���� ����������� � AccCol ���� � SF
' ��� Case Sensitive ��� ���� �����������, � ������� �� VLOOKUP
'   30.4.12
'  1.5.12 - �������� �� ������ ���� ������ � ��������
'  3.5.12 - ������������� Hash ������ �.����

    Dim L, i, j, AccCol As Integer
    Dim Acc, MSG As String
    Dim accInSF As String
    Dim t0, t1
    
    Static accHTable(0 To 5000) As String
    Static accVTable(0 To 5000) As String
    
    hashInit accHTable, accVTable
    
    t0 = Timer
    
    MSG = "SFaccColFill -- ���������� ������� 1 ��� ����� " & SheetN
    
    ModStart SheetN, MSG, True
    
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
            MsgBox "������������ ���� � SFaccColFill", , "FATAL ERROR"
            Stop
    End Select
    
 ' -- ������ �� ����� ����������� � SF SFacc ��� ������ ������ SheetN
    With Sheets(SFacc)
        For i = 2 To L
            Progress i / L
 
            Acc = Replace(Compressor(Sheets(SheetN).Cells(i, AccCol)), vbCrLf, "")
            If Acc = "" Then
                accInSF = ""
            Else
                accInSF = hashGet(accHTable, accVTable, Acc)
            End If
            If accInSF = "$" Then
                Sheets(SheetN).Cells(i, 1) = ""
                If Acc <> "" Then
                    For j = 2 To EOL_SFacc
                        If Acc = Replace(Sheets(SFacc).Cells(j, SFACC_ACC1C_COL), "" & vbCrLf, "") Then
                            Sheets(SheetN).Cells(i, 1) = "1"
                            Exit For
                        End If
                    Next j
                End If
                hashSet accHTable, accVTable, Acc, Sheets(SheetN).Cells(i, 1)
            Else
                Sheets(SheetN).Cells(i, 1) = accInSF
            End If
        Next i
    End With

    t1 = Timer - t0
'    MsgBox "Time =" & t1

    ModEnd SheetN
End Sub
Sub SFaccCol(SheetN, Optional Reslines As Integer = 0)
'
' ������������ ������� �:� � ����� SheetsN �� SFacc - ����������� �����������
' ���� ������ ����������� ���� � SF (�������� � ������� =1)��� ������, ����� �������
'   12.1.2012
'   25.1.2012 - ���������� ������ � Lines; ������ ������������ Private L
'   29.2.2012 - bug fix, EOL ������ AutoFulterReset
'   13.5.2012 - ����� ������� ��������� �����

    Dim L
    
    L = EOL(SheetN) - Reslines
    
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
