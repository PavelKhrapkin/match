Attribute VB_Name = "ADSKPartnerCenterPass"
'----------------------------------------------------------------------------
' ������ ADSKPartnerCenterPass - ������ � License Inquiry PartnerCenter.Autodesk.com
'
''' ���� ���� ��������� � �������� ��������� SN3PASS:
''' 1)                - ������� ���� 3PASS
''' 2) SNselection    - ������� ������ SN �� ����� ADSKfrSF
''' 3)   **           - �� ������ �� ����� 3PASS ��������� (Copy/Paste) SN
'''                     �������� �� ������ ADSK � Tab <License Inquiry>, ����� <Go>
'''                     ��� ������������ ��������� �� ������ [2] �����������
'''                     �� ���������� ������ SN � �2 - ������ ������, � �3 - �����
''' 4) Load3passSN    - ��������� �� ���� 3PASS �������� ����� �� Autodesk
''' 5) DoDeDupSN      - ������������ �������� ������� �� ����� 3PASS
''' 6) WrDL3pass      - ������ �� ������� 3PASS � ���� CSV ��� Data Loader
'
'   8.2.2012
'  13.8.2013 - ���������� ��� match 2.0

    Option Explicit     ' Force explicit variable declaration
    
    Dim Stage               ' ������ �������� 3PASS
    Private N As Integer    ' ����� ������ ������ � SN �� ����� 3PASS
    
Sub SN_PC_pass()
'
' S SN_PC_Pass - ��������� �������� ������� (SN) �� PartnerCenter.Autodesk.com
'                �� ������ � SN_LIST �� �������, � ����� ���������� ���� ������
'                �� SN_ACTIVE � SN_UPDATE.
'
'   13.8.2013

    Const SN_LIST = "A_PC_1"
    Const SN_TMP = "A_PC_2"
    Const ST_ACTIVE = "A_PC_3"
    Const ST_UPDATE = "A_PC_4"
    
    Dim R As TOCmatch
    Dim i, N As Integer ' ������ ������, ������
    Dim FrN, ToN, LastN ' ������ � ����� ������ SN, ��������� SN
    Const HelpStage = " ����������� � ����� ��������." & vbCrLf & _
        vbCrLf & "������:" & _
        vbCrLf & "1. � ���� Serial Number <License Inquiry>" & _
        vbCrLf & "   PartnerCenter.Autodesk.com ����� Ctrl/V, ����� [Go]" & _
        vbCrLf & "2. � Menu -> Export -> Next ������ ����� output.csv." & _
        vbCrLf & _
        vbCrLf & "����� ��������� output.csv � �������� <��������> ��� [OK]."
                        
    Const SNstep = 100  ' ������������ ������ SN
    
    StepIn
    R = GetRep(SN_LIST)
    MS "��������� � PartnerCenter.Autodesk.com " & R.EOL & " ����� Registered SN"
    
 '---------- ������������� -------------------
    iActive = 1: iUpdate = 1
    nActive = 0: nUpdate = 0: nToClean = 0
    
    Dim Rtmp As TOCmatch, Ractive As TOCmatch, Rupdate As TOCmatch
    
    Rtmp = GetRep(SNtmp)
    Ractive = GetRep(SNactive)
    Rupdate = GetRep(SNupdate)
    NewSheet Ractive.Name
    NewSheet Rupdate.Name
    
  '--------------- ���� �� SN_List ---------------
    i = 2
    With Workbooks(R.RepFile).Sheets(R.SheetN)
        .Activate
        Do While i <= R.EOL
            ToN = i + SNstep
            If ToN > R.EOL Then ToN = R.EOL
            .Cells(ToN, 1) = Left(.Cells(ToN, 1), 12)   ' ������� ��������� +
            Range(.Cells(i, 1), .Cells(ToN, 1)).Select  ' �������� ������
            Selection.Copy
            Selection.Interior.Color = rgbYellow        ' ���������� ������ ������
            
            On Error Resume Next
            Kill DownloadDir & "output*.csv"    ' ������� ��� ������ ����� output*.csv
            On Error GoTo 0                     ' �� ������, ����� output*.csv �����������
            
            MsgBox "SN �� " & i & " �� " & ToN & HelpStage
            SNread SN_TMP
            SNsortOut SN_TMP, SN_ACTIVE, SN_UPDATE
        Loop
    End With
 End Sub
Sub SNsortOut(SNtmp, SNactive, SNupdate)
'
' - SNsortOut(SNtmp, SNactive, SNupdate) - ������ � �������� SN �� SNtmp.
'      - � SNactive ��������� ������ � Registered SN
'      - � SNudpade - �� Upgraded SN, ������� ���� �������� � SF
'   ����� ���� � SNtmp ���������� ���� SN:
'           - �����      - SN ������, �� Registered
'           - ���������� - ������, ��������� Update
'           - �������� ������ - �� ������ ��� ������� ������ ��������
'
'   13.8.2013

    Dim iTmp As Long    '- ��������� - ����� ������ � SNtmp
    Call SheetSort(Rtmp.Name, 1)
    
    For iTmp = 2 To Rtmp.EOL
    
    Next iTmp
End Sub

Sub LoadSNfrOutputCSV()
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
    F = dir(DownloadDir & "output*.csv")
    i = 2
    Do While F <> ""
        Cells(i, 2) = DownloadDir & F
        Cells(i, 2).Select
        F = dir()
        i = i + 1
    Loop

' ������ ������ �� ������ � �
    R = 5: i = 2: C = 7
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

    ' ���������, ��� �� SN �� ������ � ������� � �������?
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

    Shell "quota_3PASS.bat"
'    Shell "Copy 3PASS.csv C:\SFconstr"

    End3PASS (4)
End Sub
Function Start3PASS(Msg)
'
' ������ �������� 3PASS. ���������� ����� ��������� �������� ������ ADSKfrSF
'   2.2.2012
    Const txt = "��������� 3PASS: ��������� SN �� PartnerCenter --  "
    Const A3PASSstamp = "SNfromSF"          ' ����� ����� 3PASS
    Const ResLines = 6  ' ����� ����� � ����� ������ ADSKfrSF
    

    ModStart A3PASS, txt & Msg
    Start3PASS = EOL(ADSKfrSF) - ResLines
    CheckSheet ADSKfrSF, Start3PASS + 2, 1, ADSKfrSFstamp
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
Sub SNread(F)
'
' - SNread  - ������� ������ ����� �������� ������� �� Output.CSV
'             � ������� ����� ���� ����� F
'   31/1/2012
'   13.08.13 - ������� ��������� - ������ ������ � A_SN_2

    Dim R As TOCmatch

    NewSheet F      ' ���������� ���� SN �� PartnerCenter
    
    R = GetRep(F)
    If R.EOL <> 1 Then ErrMsg FATAL_ERR, "���� \\W_TMP.A_PC_2' �� ������!"
    Workbooks(R.RepFile).Sheets(R.SheetN).Activate
    
    With ActiveSheet.QueryTables.Add( _
            Connection:="TEXT;" & DownloadDir & "output.csv", _
            Destination:=Cells(1, 1))
        .Name = "output"
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
        On Error GoTo Rep
        .Refresh BackgroundQuery:=False
    End With
    Exit Sub
Rep:
    If MsgBox("� �������� '��������' �� ������ ���� 'output.csv'." _
        & vbCrLf & "'" & DownloadDir & "output.csv" & "'" _
        & vbCrLf & vbCrLf & "�������, ��� �� �������� �� PartnerCenter � �������.") _
        = vbYes Then
            Call SNread
    Else
            End
    End If
End Sub
