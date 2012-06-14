Attribute VB_Name = "ADSK3PASS"
'----------------------------------------------------------------------------
' ������ ADSK3PASS - ������ � PartnerCenter.Autodesk.com �� �������� �������
'
' �������� ������ ������ (3PASS) ��������� �����������
' ���� ���� ��������� � �������� ��������� SN3PASS:
' 1)                - ������� ���� 3PASS
' 2) SNselection    - ������� ������ SN �� ����� ADSKfrSF
' 3)   **           - �� ������ �� ����� 3PASS ��������� (Copy/Paste) SN
'                     �������� �� ������ ADSK � Tab <License Inquiry>, ����� <Go>
'                     ��� ������������ ��������� �� ������ [2] �����������
'                     �� ���������� ������ SN � �2 - ������ ������, � �3 - �����
' 4) Load3passSN    - ��������� �� ���� 3PASS �������� ����� �� Autodesk
' 5) DoDeDupSN      - ������������ �������� ������� �� ����� 3PASS
' 6) WrDL3pass      - ������ �� ������� 3PASS � ���� CSV ��� Data Loader
'
'   8.2.2012

    Option Explicit     ' Force explicit variable declaration
    
'    Public Const A3PASS = "3PASS"       ' ���� ��� �������� 3PASS

    Dim Stage               ' ������ �������� 3PASS
    Private N As Integer    ' ����� ������ ������ � SN �� ����� 3PASS
    
Sub SN3pass()
'
' [*] - ������ 3PASS - �������� �������, �.�. ������ ������ �������� �������
'       � ���� License Inquiry ������� PartnerCenter �� ������ �� ������ SF.
'       ����� ��������� ������ �� PartnerCenter �������������� �� ����� 3PASS
'
'   8.2.2012

    Dim FrN, ToN, LastN ' ������ � ����� ������ SN, ��������� SN
    Dim i, N As Integer ' ������ ������, ������
    Const HelpStage2 = " ����������� � ����� ��������." & vbCrLf & _
        vbCrLf & "������:" & _
        vbCrLf & "1. � ���� Serial Number <License Inquiry>" & _
        vbCrLf & "   PartnerCenter.Autodesk.com ����� Ctrl/V, ����� [Go]" & _
        vbCrLf & "2. � Menu -> Export -> Next ������ ����� output.csv." & _
        vbCrLf & _
        vbCrLf & "����� ��������� output.csv � �������� <��������> ��� [OK]."
                        
    Const SNstep = 100  ' ������������ ������ SN
      
    Lines = Start3PASS("C�c�������� ������ SN")
    
    Rows("5:7777").Delete       ' ������� ������� ���������� ����� 3PASS
    On Error Resume Next
    Kill DownloadDir & "output*.csv"    '   .. � ����� output*.csv
    On Error GoTo 0             ' �� ������, ����� output*.csv �����������
    
    SNselectionForm.Show
    Sheets(A3PASS).Select
    FrN = Cells(2, 1)
    ToN = Cells(3, 1)
    If FrN >= ToN Then End      ' ���� �� ������ SN �� ������� - �����
     
'** �������� ������ �� PartnerCenter � ����� output.csv �������� "��������"
    Start3PASS ("���������� SN �� PartnerCenter �������� � ����� output*.csv")
    LastN = ToN
    N = Round((LastN - FrN) / SNstep) + 1

    For i = 1 To N
        If ToN - FrN > SNstep Then ToN = FrN + SNstep
        Cells(ToN, 1) = Left(Cells(ToN, 1), 12) ' ������� ��������� +
        Range("A" & FrN & ":A" & ToN).Select    ' �������� ������
        Selection.Copy
        Selection.Interior.Color = rgbYellow    ' ���������� ������ ������
        MsgBox i & "/" & N & ": SN �� " & FrN & " �� " & ToN & HelpStage2
        FrN = ToN + 1
        ToN = LastN
    Next i
    
    Load3passSN
'    DoDeDupSN
'    WrDL3pass
'    End3PASS 3
End Sub

Sub Load3passSN()
'
' [3] - ������ SN ���������� �� Licence Inquiry � ������ output.csv
' ������ ����� ���� ���������. ����� ��������� ��� ���������.
'
'   1.2.2012

    Const CSVstamp1 = "Serial Number"
    Const CSVstamp2 = "Product Key"
    Dim F, SN, Status As String
    Dim R, C, i, j As Integer
    
    Start3PASS ("�������� ������ output*.csv")
        
' ������ ������ ������ output*.csv � ������� B ����� 3PASS
    F = Dir(DownloadDir & "output*.csv")
    i = 5
    Do While F <> ""
        Cells(i, 2) = DownloadDir & F
        Cells(i, 2).Select
        F = Dir()
        i = i + 1
    Loop

' ������ ������ �� ������ � �
    R = 5: i = 5: C = 7
    Do While Cells(i, 2) <> ""
        Cells(i, 2).Select
        Call SNread(Cells(i, 2), R, C)    ' ������ ����� output.csv
    ' ��������, ���������� �� ���� ���������?
        If Cells(R, C) <> CSVstamp1 Or Cells(R, C + 1) <> CSVstamp2 Then
            MsgBox "������������ ������� ���� (" & F & _
                ") � ����� 3PASS, ������� � ������ (" _
                & R & "," & C & ")", vbCritical, "ERROR!"
            Stop
        End If
    ' �������� ����� ������������ �����
        Range("G" & R & ":V" & R).Delete Shift:=xlUp
        While Cells(R, C) <> "" ' ���������� R �� ������ ������
            R = R + 1
        Wend
        i = i + 1
    Loop

    R = R - 1
    i = 5
    Do
        SN = Left(Cells(i, 1), 12)        ' ������ SN
'        Cells(i, 1).Select
        For j = 2 To R
            If SN = "" Then Exit Do
            If SN = Cells(j, 7) Then
                i = i + 1
                Exit For
            End If
        Next j
    Loop
        
    If i - 1 <> Cells(3, 1) Then
        MsgBox "�� ��� �������� ������ ������� -- ��.���� 3PASS", , "ERROR!"
        Stop
    End If
        
    Range("B5:B" & i).ClearContents
    Range("A4:A" & Cells(3, 1)).ClearContents
    Range("B4:F" & R).FillDown  ' �������� ������� �� ����� (R)

    End3PASS 3
End Sub
Sub DoDeDupSN()
'
' [*] ������������ �������� ������� �� ����� 3PASS
'   6.2.2012

End Sub
Sub WrDL3pass()
'
' [3PASS] - ������ ��� ������ Data Loader'a ����� 3PASS � Salesforce
'   3/2/2012

    Start3PASS "������ ����� 3PASS ����������� DataLoader � SF"
    
    Columns("C:C").Copy
    Columns("O:O").PasteSpecial Paste:=xlPasteValues
'    Columns("O:O").Select
'    Selection.PasteSpecial Paste:=xlPasteValues
    
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\ADSK\3PASS\"
    WriteCSV A3PASS, "3PASS.txt", 1, 7

    Shell "quota2.bat 3PASS.TXT C:\SFconstr\3PASS.csv"
'    Shell "Copy 3PASS.csv C:\SFconstr"

    End3PASS (4)
End Sub
Function Start3PASS(Msg)
'
' ������ �������� 3PASS. ���������� ����� ��������� �������� ������ ADSKfrSF
'   2.2.2012
    Const txt = "��������� 3PASS: ��������� SN �� PartnerCenter --  "
    Const ASFstamp = "Match ADSK from SF"   ' ����� - ��� ������ SF � �����
    Const A3PASSstamp = "SNfromSF"          ' ����� ����� 3PASS
    Const Reslines = 6  ' ����� ����� � ����� ������ ADSKfrSF
    

    ModStart A3PASS, txt & Msg
    Start3PASS = EOL(ADSKfrSF) - Reslines
    CheckSheet ADSKfrSF, Start3PASS + 2, 1, ASFstamp
    CheckSheet A3PASS, 1, 1, A3PASSstamp
End Function
Sub End3PASS(M)
'
' ���������� ����������� Stage3PASS
'   30/1/2012

    Const Colr1 = &HFFCC00  ' ���� [1] - �������
    Const Colr2 = &HFF9900  ' ���� [2] - �����
    Const Colr3 = &HFF6600  ' ���� [1] - �����-�����
    Dim Colr
    
    Sheets(A3PASS).Select
    Select Case M
        Case 1
            Colr = Colr1
        Case 2
            Colr = Colr2
        Case 3
            Colr = Colr3
    End Select
    Cells(1, 1).Interior.Color = Colr   ' ���������� ������ �1
    Sheets(A3PASS).Tab.Color = Colr     '   .. � Tab 3PASS
    ModEnd A3PASS
End Sub
Sub SNread(FileNM, R, C)
'
' ������� ������ ����� �������� ������� FileNM �� ���� 3PASS � ������� (R,C)
'   31/1/2012

 '   FileNM = "C:\Users\������������\Downloads\output.csv"
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & FileNM, _
        Destination:=Cells(R, C))
'        .Name = "output"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = _
            Array(2, 2, 2, 2, 2, 2, 2, 2, 3, 2, 2, 2, 2, 2, 2, 2)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub
