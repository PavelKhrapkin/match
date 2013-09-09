Attribute VB_Name = "WPhandlings"
'====================================================================
' ������ ��� ������ � ������ ������� ������ WP
'   1.10.2012
'
' ����� ������ ������ � Match ��� �������������� �� ������� ����� WP

    Option Explicit    ' Force explicit variable declaration
Sub WP_PaidOpp(Form As String)
'
'
'

End Sub




''    Const WP = "WP"     ' ��� �������� �����
''    Const We = "We"     ' ���������� �������� � ����� ����������� � ���������
''    Const WPheader = "WPheader" ' ���� - ������ ��� WP
''
''    Dim PaidMod         ' ����� ������ ��������
''    Dim OppMod          ' ����� ������ ��������
''    Dim ContrMod        ' ����� ������ ���������
''    Dim ADSKmod         ' ����� ������ �������� Autodesk
''
''    Dim Prow            ' ������� ��������� �� ����� ��������
''    Dim Pcol            ' ����� ������� �� ����� ��������
''    Dim Wppay0, WPpayN, Wpopp0, WPoppN, Wpdod0, WPdonN
''    Dim EOL_WP
    
Sub AccPaid_Click()
'
' �� ������ <WP> �� ������ �������� ������ �� ��������� ����������
' �� ����� WP, � ���������:
'   - ���� ����������� ��� � SF - ������������� � �������� �� 1�
'   - ����� �������� ����������� (���������� � �� ���������� � SF)
'   - ����� �������� ����������� (����������� � �������������)
'   - ����� ��������� (������ ���������� � �������� � ���� - � SF � 1�)
'   - ����� � ������� ��������� ADSK (� ����������� ��������� � ����)
'
'   17/1/12
'   10/2/12 - ������������� WPform

    Dim Client1C As String
    
    Prow = ActiveCell.Row           ' ����� ��������� ������
    Pcol = ActiveCell.Column        ' ��������� �������
    Client1C = Range("I" & Prow)    ' ��������� �����������
    
    ModStart WP, "������ � WP -- ����������� � ��������� �������� �� 1�"
    CheckSheet 1, 1, 6, "����. ���."
    CheckSheet 1, 1, 7, "���� ����. �����"
    
    If Prow > Sheets(1).UsedRange.Rows.Count - 3 Or _
       Pcol > Sheets(1).UsedRange.Columns.Count Then
        MsgBox "����� �� ��������� ��������!", , "ERROR!"
        Exit Sub
    End If

    Call WPaid(-2)           ' �������� �������� �� ����������� �� WP
    
    If Sheets(1).Cells(Prow, 1) = "" Then
        AccAddFrom1C Client1C   ' ��������� ����������� ��� � SF - ���������
        Exit Sub
    End If
    
    Wopp Client1C   ' ����� �������� � ������ OppFlag
    Wcontr Client1C ' ����� ��������� � ������ ContrFlag
    WADSK Client1C  ' ����� ��������� � ����������� ��������� Autodesk � ������ ADSKFlag
    
    WPform.Show
    ModEnd WP
End Sub
Sub WPaid(Optional PaidMod)
'
' ����� �������� �� ��������� ����������� �� ����� WP
'   PaidMod<0   �������� � �������� ����� WP � ������ -PaidMod
'   PaidMod=1   �������� ������ ��������� ������
'   PaidMod=2   �������� ������ �������, ������� ���� � SF (�������)
'   PaidMod=3   �������� ��� ������� �� ��������� �����������
' �� ����� WP ���� ��� ������� (2), �� �������� ��������
'
'   19.1.2012
'    4.2.2012 - Optional Public PaidMod. ���� �������� ������ -
'               �������������� Public PaidMod
    
    Static FrL, ToL As Integer  ' �������� ������� ����� � ��������� �� ����� WP
    Dim P_Lines As Integer      ' ����� �����- �������� ����� 1 ��� �����
    Dim i As Integer
    Dim Index As String
    
    If PaidMod < 0 Then
        ClearSheet WP, Range("HDR_WP")
        FrL = EOL(WP) + 1                       ' ��������� ������ �������� �� WP
        P_Lines = AutoFilterReset(1) - 3
        Range("A1:AC" & EOL(1) - 3).AutoFilter Field:=9, _
            Criteria1:="=" & Range("I" & Prow)  ' ������ � 1 �� ��������� �����������
        Give "A2:AC" & P_Lines, "A2", 1  ' ��� ������� ����������� � WP
        PaidMod = -PaidMod
        ToL = EOL(WP)
    End If
       
    Select Case PaidMod
        Case 1                      ' ����� 1 - ������� ������ ��������� ������
            HideLns FrL, ToL, 0                 ' Hide ���
            Index = Sheets(1).Cells(Prow, 2)
            HideLns FrL, ToL, 2, Index, False   ' Unhide ��������� ������
        Case 2                      ' ����� 2 - ������� ������ �������, ������� ��� � SF
            HideLns FrL, ToL, 0                 ' Hide ���
            HideLns FrL, ToL, 4, , False        ' Unhide �������, ������� ��� � SF
        Case 3                      ' ����� 3 - ��� ������� ��������� �����������
            HideLns FrL, ToL, 0, , False        ' Unhide ���
        Case Else
            MsgBox "ERROR in WPaid: PaidMod = " & PaidMod, , "ERROR!"
    End Select
'    If Prow <= 4 Or Prow > Lines Then Prow = Lines
'
'    Prow = ActiveCell.Row                               ' ������� ������ (Prow,Pcol)
'    Pcol = ActiveCell.Column
'    If Prow > 4 And Prow <= Lines And Pcol < 17 Then    ' ���� ��� � ���������� ���������
'        Range(Cells(Prow, 6), Cells(Prow, 17)).Select
'        Selection.Interior.Color = rgbLightCoral        ' �������� ������
'        Range(Cells(Prow, Pcol)).Interior.Color = rgbCoral
'        Selection.Interior.Color = rgbCoral             '   .. � ����� ���� - �� ����
'    Else
'        Prow = Lines: Pcol = 6
'    End If
End Sub
Sub WP_Paid_Click()
'
' ������������ ��������� ���������� �� ������ �������� �� ������� �����
'   19.1.12
    
    PaidMod = PaidMod + 1
    If PaidMod = 3 Then PaidMod = 0
    Call WPaid
'
'    Call Wopp
'    Call Wcontr     ' ����� ��������� � ������ ContrFlag
'    Call WADSK      ' ����� ��������� � ����������� ��������� Autodesk � ������ ADSKFlag

End Sub
Sub Wopp(Account)
'
' ����� ������ �������� �� ����� WP ��� ��������� ����������� (Account)
'   25/1/2012
'   10.2.2012 - �������

    Dim OppLines, WPlines, i As Integer
    
    Lines = EOL(WP) + 2                     ' ��������� ������ ��� �������� �� WP
    AutoFilterReset SFopp
    OppLines = EOL(SFopp) - SFresLines      ' ����� ����� � SFopp
    Range("A1:D" & OppLines).AutoFilter Field:=4, _
        Criteria1:="=" & Account            ' ������ SFopp �� ��������� �����������
   
    Give "B1:B" & OppLines, "B" & Lines, 2  ' ����� �������
    WPlines = EOL(WP)                       ' ����� ����� WP ������ � ��������
    ActiveSheet.Range("B1:B" & WPlines).HorizontalAlignment = xlCenter
    Give "C1:C" & OppLines, "K" & Lines, 2  ' ��� �������
    Give "H1:I" & OppLines, "G" & Lines, 2  ' ���� �������� � �������� �������
    Give "J1:J" & OppLines, "L" & Lines, 2  ' ��������
    Give "O1:O" & OppLines, "J" & Lines, 2  ' ��� �� ��������� (���/USD/EUR)
    Give "P1:P" & OppLines, "P" & Lines, 2  ' �� �������� %
    Give "M1:M" & OppLines, "N" & Lines, 2  ' ���������
    Give "G1:G" & OppLines, "D" & Lines, 2  ' ���� "�����������"
    Give "E1:E" & OppLines, "F" & Lines, 2  ' ����������� (%)
    Give "Q1:Q" & OppLines, "I" & Lines, 2  ' ����&��� ������������
    Give "N1:N" & OppLines, "A" & Lines, 2  ' ������ �������

'**** ������ �� ����������� ��������
    For i = Lines + 1 To WPlines        ' ����������� ������ �������� �� WP
        PerCent i, 6                    ' ����������� %
        PerCent i, 16                   ' �� �������� %
        CurCode i, 10, 1                ' ������ ������ ���������
        CurCode i, 14, 1                ' ������ ������ "��� �� ���������"
'        If Cells(i,6) = "100%  ' ���������� ���������� ������
    Next i
    
    Range("F1").Select                  ' �������� ������ ��������� �� ��������
    Selection.Copy
    Range("B" & Lines & ":Q" & Lines).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Rows(Lines & ":" & WPlines).Font.Name = "Calibri"
    Rows(Lines & ":" & WPlines).Font.size = 8
    Range("A" & Lines + 1 & ":A" & WPlines).WrapText = False ' ��������� ������ �������
    Range("A" & Lines & "," & "D" & Lines).Clear
    
    HideLns Lines, WPlines, 6, "100%"

End Sub
Sub Wcontr(Account)
' ����� ������ ��������� �� ����� WP ��� ��������� ����������� (Account'a)
'   25/1/2012
' 10.2.2012 - �������
   
    Dim WPlines, ContrLines, Index, i
   
    Lines = EOL(WP) + 2                     ' ��������� ������ ��� ��������� �� WP
    AutoFilterReset SFD                       ' ���������� ��� ������� � ������� �� 3
'    SheetSort 3, 4                          ' ��������� �������� �� ������� � 1�
    ContrLines = EOL(SFD) - SFresLines      ' ����� ����� � SFD
    Range("A1:F" & ContrLines).AutoFilter Field:=6, _
        Criteria1:="=" & Account            ' ������ SFD �� ��������� �����������
    
    Give "A2:A" & ContrLines, "B" & Lines + 1, 3   ' ������ �������� (������ �� ���� 3)
    WPlines = EOL(WP)                       ' ����� ����� WP ������ � ����������
    ActiveSheet.Range("B1:B" & WPlines).HorizontalAlignment = xlCenter
    Give "M1:M" & ContrLines, "F" & Lines, 3   ' ������ ��������
    Give "D1:E" & ContrLines, "M" & Lines, 3  ' ������� � �������� �������
    Give "G1:G" & ContrLines, "I" & Lines, 3  ' ��� ��������
    Give "H1:I" & ContrLines, "G" & Lines, 3  ' ���� ������ � ���������� ��������
    Give "P1:P" & ContrLines, "K" & Lines, 3  ' ������
    Give "Q1:Q" & ContrLines, "P" & Lines, 3  ' ����� �������

'**** ������ �� ����������� ���������
    Sheets(WP).Select
    For i = Lines + 1 To WPlines
        Index = Cells(i, 2)     ' ������ ������� - ����� �� ����� ��������� �� 1�
        Sheets(3).Range("E" & Index + 1).Copy
        Cells(i, 1).Select
        ActiveSheet.Paste
'        PerCent i, 6                    ' ����������� %
'        PerCent i, 16                   ' �� �������� %
'        CurCode i, 10, 1                ' ������ ������ ���������
'        CurCode i, 14, 1                ' ������ ������ "��� �� ���������"
'        If Cells(i,6) = "100%  ' ���������� ���������� ������
    Next i
    
    Range("F1").Select                  ' �������� ������ ��������� �� ��������
    Selection.Copy
    Range("B" & Lines & ":Q" & Lines).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    Rows(Lines & ":" & WPlines).Font.Name = "Calibri"    ' ������ �����
    Rows(Lines & ":" & WPlines).Font.size = 8
    Range("A" & Lines & ":A" & WPlines).WrapText = False
    Range("A" & Lines & "," & "D" & Lines).Clear    ' ������� ������� ��������� �������
    
    HideLns Lines, WPlines, 6, "������"
    HideLns Lines, WPlines, 6, "��� � SF"

End Sub
Sub WADSK(Account)
' ����� ������ ��������� Autodesk �� ����� WP ��� ��������� �����������
'   12.2.2012

End Sub

Sub AccAddFrom1C(Acc As String)
'
' ������� ������� � ���������� ����������� Acc �� 1� � SF
'   5.2.2012

'    Dim Row1C   ' ������ - ����� ������ ����������� � 1�
    Dim adr As String
    Dim AdrLength As Integer
    Dim AdrArr() As String
    
    Dim PostIndex As Integer
    Dim City, Street, Region, Country As String
'    Dim Q
    
    
    ModStart Acc1C, "����� ����������� " & Acc & " �� 1C"
    
    adr = Adr1C(Acc)
    adr = Trim(adr)
    AdrLength = Len(adr)
    
    AddressForm.TextBox1 = Acc
    AddressForm.TextBox2 = adr
    
    AdrArr = split(adr, ",")
    AddressForm.PostIndex = AdrArr(0)
    AddressForm.City = AdrArr(1)
    AddressForm.Street = AdrArr(2)
    AddressForm.PostCountry = "������"
    
    AddressForm.Show
    
  '  PostIndex = Left(Adr, 6)
    MsgBox "����� � ����������� 1�=" & adr & " ������=" & PostIndex
        
    NewSFaccForm.Show
'    Do
'        AccSF = SearchSFacc(Acc)
'        If AccSF <> 0 Then
'            Q = MsgBox("��� ������������� " & AccSF & " � SF?", vbYesNoCancel)
'            If Q = vbYes Or Q = vbCancel Then
'                Exit Do
'            End If
'        End If
'    ' ����������� ������ - ����� ���
''?????????????????????????????????????????????
'    Loop
        
    ModEnd WP
End Sub
Sub SearchSFacc(txt)
'
' ����� � ������ ����������� SFacc ��������� txt �� ����� �����������
' ���������� ����� ������ � SFacc ��� 0, ���� �� �������
'   3.2.2012

    Dim Row, i As Integer
    Dim S, Str As String

  
    For Row = 2 To Lines
        Str = Sheets(SFacc).Cells(Row, 1)
S = Application.WorksheetFunction.Search(txt, Str)
'        i = Search(txt, Str)
        MsgBox i
    Next Row
End Sub
Sub Give(Fr1, ToWP, Mode)
'
' ������������ ����������� ������ Fr1 �� ����� ToWP �� ������� ���� WP
' ����  Mode = 1 - ����������� PasteSpecial �� ��������
'       Mode = 2 - �������� �� SFopp
'       Mode = 3 - �������� �� SFD
' 22.1.2012

    Select Case Mode
        Case 1
            Sheets(1).Range(Fr1).Copy
            Sheets(WP).Select
            Range(ToWP).Select
            Selection.PasteSpecial Paste:=xlPasteColumnWidths, _
                Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Case 2
            Sheets("SFopp").Range(Fr1).Copy
            Sheets(WP).Select
            Range(ToWP).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
        Case 3
            Sheets("SFD").Range(Fr1).Copy
            Sheets(WP).Select
            Range(ToWP).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
    End Select

    ActiveSheet.Paste
End Sub
