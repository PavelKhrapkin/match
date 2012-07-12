Attribute VB_Name = "FromSF"
'---------------------------------------------------------------------------
' ���������� ������� �� salesforce.com � Match SF-1C.xlms
'
' ������������ � ������� (*):
' (*) ShowDBGbanel - ����� ������ ��� ������� ��������� �� Ctrl/Shift/Q
' [*] Match1C-SF()- ��������� ���� SF- ������� �� �������� ��� ������ SF � 1�
' [*] SFDreport() - �������� ���� SFD ����� ������� �� ��������� �� Salesforce
' [*] SFaccRep()  - �������� ���� SFacc ����� ������� �� ������������ �� SF
' [*] SFoppRep()  - �������� ���� SFopp ����� ������� �� �������� �� SF
' [*] ADSKfromSF()- �������� ����� �� SF �� Autodesk "Match ADSK from SF" �����
' (*) SF_PA_Rep   - p����� ������ SF_PA - ������ �������� � �����.ADSK �� SF
'
' �.�.������� 2.1.2012
'   5.2.2012 - ������ �� �������������� ���� � ������� ������, �������
'   8.2.2012 - �������� (CheckSheet) "����������� ����� �����?"
'  29.2.2012 - �� �������� ����� � ����� ����� SFopp
'   4.3.12 - bug fix
'  30.3.12 - bug fix EOL � SFD
'  28.4.12 - �������� ���� SForders
'  30.4.12 - SFaccColFill
'  12.5.12 - bug fix
'  16.5.12 - ����� ����� �� ������ �������� � ����������� ADSK SF_PA

    Option Explicit    ' Force explicit variable declaration
        
Sub ShowControlPanel()
Attribute ShowControlPanel.VB_Description = "8.2.2012 ������ ShowDBGpanel "
Attribute ShowControlPanel.VB_ProcData.VB_Invoke_Func = "Q\n14"
'
' ����� ���������� ����������� ������ � ���������� �������� �� ���� ��������� ������
'
' Ctrl/Shift/Q
'
'   8.2.2012

    MainControlPanel.Show
    End
End Sub
Sub Match1C_SF()
'
' (*) ���������� ������ SF "������� ������ � 1�"
'
' ����������� �������� [SF] �� ����� �������� � "SF"
'
'   4.9.11
'   5.1.12 - ������������� ������������ PaintPayment, Application.StatusBar
'   9.1.12 - ���������� ����������� ������ �� SF
'  26.1.12 - ��������, ��� �� ����� ������������� ����� �������, ���������� SF
'  28.1.12 - �������������� �� ������ ������
    
    LinesOld = ModStart("SF", _
        "MatchSF_1C - ��������� ���� SF �� ������ Salesforce <������� �� 1�>")
    
    Lines = Sheets(1).UsedRange.Rows.count      ' ���-�� ����� � ����� ������
    CheckSheet 1, Lines - 4, 1, SFpayRepName    ' ��������� ������������ ������ �
    CheckSheet "SF", LinesOld + 2, 3, SFpayRepName  '   .. �������� ������� SF
    SheetSort 1, 3                      ' ����� ����� ��������� �� �����
    
    Sheets("SF").Select
    Columns("A:B").Copy                 ' �������� 2 ������ �������
    
    Sheets("SF").Name = "RRR"           ' ������� ����� SF ��������������� � RRR
    
    Worksheets(1).Name = "SF"           ' ����� ����� ��������������� � SF
    Sheets("SF").Select
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight   ' ��������� 2 ������� �� �������� �����

    Rows("2:" & Lines).RowHeight = 15
    
    Sheets("RRR").Range("E" & LinesOld - 1 & ":G" & LinesOld).Copy
    Sheets("SF").Range("E" & Lines - 1).Select
    ActiveSheet.Paste                   ' �������� ������ CSIT
    Sheets("RRR").Range("A" & LinesOld - 1 & ":B" & LinesOld).Copy
    Sheets("SF").Range("A" & Lines - 1).Select
    ActiveSheet.Paste   ' �������� ������� ������ - "�� ������� � ��������� 1�"

' ���� ����, ���������� ���� ������� ������
    If Lines > LinesOld Then Range("A" & LinesOld - SFresLines & _
                                  ":B" & Lines - SFresLines).FillDown
    
    With Range("A" & Lines - 5 & ":B" & Lines - 2)
        .ClearContents
        .Interior.Color = rgbWhite      ' ������� ��� �������� �� ����� ������
    End With

    Call SheetsCtrlH(2, "RRR!", "SF!")  ' �������� ������ � �������� �������� 1�
'    Call SheetsCtrlH("P_Update", "RRR!", "SF!") '   � DL Payment_Update
    
    Sheets("SF").Move After:=Sheets(2)  ' ��������� ����� ����� SF �� ������ �����
    Sheets("RRR").Select
    ActiveWindow.SelectedSheets.Delete  ' ���������� ������ ����� SF
    
'********************
    PaymentPaint   '*
'********************

    Sheets("SF").Tab.Color = rgbViolet ' ���������� Tab ������ ������ SF
    ModEnd 1
End Sub
Sub SFDreport()
'
' 3. �������� ���� SFD ����� ������� �� ��������� �� SF
'
'   6.1.2012 - ��� ������ SFD CTRL/H � P_ ������; ������� A � � - ���� � �������� SFD
'  26.1.2012 - �������� ������������ ����� CheckSheet
'   4.3.12 - bug fix FillDown
'  30.3.12 - bux fix -- EOL ������ RowsCount
    
    Dim LO, Ln As Integer           ' ���������� ����� � ������ � ����� �������
    Dim OOO, NNN
    
    ModStart "SFD", "���������� ����� SFD - ������ Salesforce �� ��������� <SFD>"

    LinesOld = EOL(SFD)       ' ���-�� ����� � ������ ������
    Lines = EOL(1)            ' ���-�� ����� � ����� ������
    LO = LinesOld - SFresLines
    Ln = Lines - SFresLines
    
    CheckSheet 1, Lines - 4, 1, SFcontrRepName
    CheckSheet "SFD", LinesOld - 4, 3, SFcontrRepName
    
    AutoFilterReset 4           ' ���������� ������� � ������ 1� �� ���������
    AutoFilterReset "SFD"       '     .. � � ����� SFD
    SheetSort 4, 4              ' ��������� �������� �� ������� ������� � 1�
    
    Sheets("SFD").Select
    Sheets("SFD").Name = "RRR"                      '��������������� SFD � RRR
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst

    Sheets(1).Name = "SFD"                          '��������������� ����� ����� � SFD
    Sheets("SFD").Select
    Sheets("SFD").Move Before:=Sheets(5)            '� ���������� ��� ����� RRR

    Sheets("RRR").Columns("A:B").Copy               ' �������� 2 ������ �������
    Sheets("SFD").Select
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight       ' ��������� 2 ������� �� ��������

    Rows("2:" & Lines).Select
    Selection.RowHeight = 15
    
    Worksheets("RRR").Range("A" & LinesOld).Copy
    Worksheets("SFD").Range("A" & Lines).Select
    ActiveSheet.Paste       ' �������� ������� ������ - �� ��������� ��������
    
    OOO = Worksheets("RRR").Range("A" & LinesOld)
    NNN = Worksheets("SFD").Range("A" & Lines)
    
    MsgBox "������� " & NNN & " �� ��������� ���������, ���� " & OOO
    
    If Ln > LO Then
        Range("A" & LO - 1 & ":B" & Ln).Select
        Selection.FillDown      ' ��������� ������� ������ � ����� ������ �� �����
    End If
       
    With Range("A" & Lines - 5 & ":B" & Lines - 1)
        .ClearContents
        .Interior.Color = rgbWhite  ' ������� ��� �������� �� ����� ������
    End With

    SheetsCtrlH 3, "RRR!", "SFD!"   ' ������ ������ � �������� ������� �� SFD
    SheetsCtrlH "P_PaidContract", "RRR!", "SFD!"
    
    Sheets("RRR").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("SFD").Select
    ActiveWorkbook.Sheets("SFD").Tab.Color = rgbViolet

'********************
    ContractPaint  '*
'********************
    
    ModEnd 3
End Sub
Sub SFaccRep()
'
' 4. ������ ������ SFacc �� ������������ �� ������� ����� ����� MatchSF-1C
'   3.1.12
'  9.2.2012 - �������
' 30.4.12 - SFaccColFill ��� ��������, ��������� � �����������
    
    Dim LO, Ln As Integer           ' ���������� ����� � ������ � ����� �������
    Dim Same As String
    
    ModStart SFacc, "���������� ����� ������ Salesforce �� ������������ SFacc"

    LinesOld = Sheets(SFacc).UsedRange.Rows.count ' ���-�� ����� � ������ ������
    Lines = Sheets(1).UsedRange.Rows.count        ' ���-�� ����� � ����� ������
    LO = LinesOld - SFresLines
    Ln = Lines - SFresLines
    
    CheckSheet 1, Lines - 4, 1, SFaccRepName
    CheckSheet SFacc, LinesOld - 4, 1, SFaccRepName
    
    If LO = Ln Then Same = "���� " Else Same = ""
    MsgBox "� ����� ������ " & Ln - 1 & " �����������, � ������� " & Same & LO - 1
                                                            
    Sheets(SFacc).Rows("1:1").Copy Sheets(1).Rows("1:1")    ' �������� ���������
    AutoFilterReset 1

    Sheets(SFacc).Name = "Tmp"          ' ��������������� ������� ����� � "Tmp"
    Sheets(1).Name = SFacc              '   .. � ����� � SFacc

' �������� ������ � ��SheetNm������ ������ ������
    Call SheetsCtrlH(2, "Tmp!", "SFacc!")
    Call SFaccCol(2, 3)                ' ���� ��������
   
    Call SheetsCtrlH(4, "Tmp!", "SFacc!")
    Call SFaccCol(4, 7)                ' ���� ���������

    Call SheetsCtrlH("DogovorHeader", "Tmp!", "SFacc!") ' ������ Dogovor_Insert

    Call SheetsCtrlH("������ �������� 1C", "Tmp!", "SFacc!")
    Call SFaccCol("������ �������� 1C", 3) ' ����� �� ����������� �������� 1�
   
    Sheets(1).Move Before:=Sheets("Tmp")    ' ���������� ����� ����� ����� �������
    Sheets("Tmp").Delete                    ' ���������� ������� ����� SF
    Sheets(SFacc).Tab.Color = rgbYellow     ' ���������� Tab ������ ������
    
    SFaccColFill PAY_SHEET  ' � ������� 1 ���� ����������� ���� � SF
    SFaccCol PAY_SHEET, PAY_RESLINES    ' ������������ ������� A
    SFaccColFill DOG_SHEET  ' � ������� 1 ���� ����������� ���� � SF
    SFaccCol DOG_SHEET, DOGRES    ' ������������ ������� A
    SFaccColFill Acc1C              ' � ������� 1 ���� ����������� ���� � SF
    SFaccCol Acc1C, ACC1C_RES       ' ������������ ������� A

    ModEnd SFacc
End Sub
Sub SFoppRep()
'
' (*) ������ ������ SFopp �� �������� �� ������� ����� ����� MatchSF-1C
'  9.2.2012
'   29.2.12 - �� �������� ����� � ����� ����� SFopp
    
    Dim LO, Ln As Integer           ' ���������� ����� � ������ � ����� �������
    Dim Same As String
    
    ModStart SFopp, "���������� ����� ������ Salesforce �� ��������� SFopp"

    LinesOld = Sheets(SFopp).UsedRange.Rows.count   ' ���-�� ����� � ������ ������
    Lines = Sheets(1).UsedRange.Rows.count          ' ���-�� ����� � ����� ������
    LO = LinesOld - SFresLines
    Ln = Lines - SFresLines
    
    CheckSheet 1, Lines - 4, 1, SFoppRepName
    CheckSheet SFopp, LinesOld - 4, 1, SFoppRepName
                                                            
'    Sheets(SFopp).Rows("1:1").Copy Sheets(1).Rows("1:1")    ' �������� ���������
    AutoFilterReset 1

    Sheets(SFopp).Name = "Tmp"          ' ��������������� ������� ����� � "Tmp"
    Sheets(1).Name = SFopp              '   .. � ����� � SFopp

' �������� ������ � �������� ������ ������

    Sheets(SFopp).Move Before:=Sheets("Tmp")  ' ���������� ����� ����� ����� �������
    Sheets("Tmp").Delete                   ' ���������� ������� ����� SF
    Sheets(SFopp).Tab.Color = rgbViolet  ' ���������� Tab ������ ������
    
    EOL_SFopp = Ln

    ModEnd SFopp
End Sub
Sub SFcontactRep()
'
' <*> SFcontactRep() -�������� ������� SFcont �� ��������� �� SF
'   9/7/12

    ModStart
End Sub
Sub ADSKfromSFrep()
'
' (*) �������� ���� ADSKfrSF ����� ������� �� Salesforce
'
' ������ [SF] � ����� ������ ������ ADSKfrSF � PartnerCenter
'
'   8.2.2012
' 12.5.12 - ����� ����������� ����� �� ������� ������ � �����
    
    Dim LO, Ln As Integer   ' ���������� ����� � ������ � ����� �������
    Dim Same As String
    
    ModStart ADSKfrSF, "���������� ����� ������ Salesforce �� Autodesk"

    LO = EOL(ADSKfrSF) - SFresLines
    Ln = EOL(1) - SFresLines
    
    CheckSheet 1, Ln + 2, 1, SFadskRepName          ' ����� ����� �����?
    CheckSheet ADSKfrSF, LO + 2, 1, SFadskRepName   ' ������� ����� �� �����?
    
    If LO = Ln Then Same = "���� " Else Same = ""
    MsgBox "� ����� ������ " & Ln - 1 & " SN, � ������� " & Same & LO - 1
    
' ���������� ����� ����� ����� ������� � ��������������� ���
    Sheets(ADSKfrSF).Name = "Tmp"
    Sheets(1).Name = ADSKfrSF
    Sheets(ADSKfrSF).Move Before:=Sheets("Tmp")
        
' �������� ������ � �������� ������ ������
'    Call SheetsCtrlH(PartnerCenter, "Tmp", ADSKfrSF)
    Call SheetsCtrlH("3PASS", "Tmp", ADSKfrSF)

    Sheets("Tmp").Delete                    ' ���������� ������� ����� SF
    Sheets(ADSKfrSF).Tab.Color = rgbViolet  ' ���������� Tab ������ ������

    ModEnd ADSKfrSF
End Sub
Sub SF_PA_Rep()
'
' (*) ������ ������ SF_PA - ������ �������� � ����������� ADSK �� SF
'   16.6.12
    
    Dim LO As Integer, Ln As Integer    ' ���-�� ����� � ������ � ����� �������
    Dim Msg As String
    
    LO = ModStart(SF_PA, "���������� ����� ������ SF_PA �� SF")
    Ln = EOL(1)

    CheckSheet 1, Ln + 2, 1, SFpaRepName
    CheckSheet SF_PA, LO + 2, 1, SFpaRepName
                                                            
    Sheets(SF_PA).Name = "Tmp"          ' ��������������� ������� ����� � "Tmp"
    Sheets(1).Name = SF_PA              '   .. � ����� � SF_PA

    Sheets(SF_PA).Move Before:=Sheets("Tmp")  ' ���������� ����� �����
    Sheets("Tmp").Delete                ' ���������� ������� ����� SF
    Sheets(SF_PA).Tab.Color = rgbViolet ' ���������� Tab ������ ������
    
    Msg = "� ������� ������ SF_PA " & LO & " �����, � ����� "
    If LO = Ln Then Msg = Msg & "���� "
    MsgBox Msg & Ln
    
    ModEnd SF_PA
End Sub

