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
'  16.8.2013 - ���������� ��� match 2.0

    Option Explicit     ' Force explicit variable declaration
    
    
Sub SN_PC_pass()
'
' S SN_PC_Pass - ��������� �������� ������� (SN) �� PartnerCenter.Autodesk.com
'                �� ������ � SN_LIST �� �������, � ����� ���������� ���� ������
'                �� SN_ACTIVE � SN_UPDATE.
'
'   16.8.2013

    Const SN_LIST = "A_SN_1"
    Const SN_TMP = "A_SN_2"
    Const SN_UPDATE = "A_SN_3"
    
    Const HelpStage = " ����������� � ����� ��������." & vbCrLf & _
        vbCrLf & "������:" & _
        vbCrLf & "1. � ���� Serial Number <License Inquiry>" & _
        vbCrLf & "   PartnerCenter.Autodesk.com ����� Ctrl/V, ����� [Go]" & _
        vbCrLf & "2. � Menu -> Export -> Next ������ ����� output.csv." & _
        vbCrLf & _
        vbCrLf & "����� ��������� output.csv � �������� <��������> ��� [OK]."
                        
    Const SNstep = 100  ' ������������ ������ SN
    
    Dim R As TOCmatch
    Dim i As Long, ToN As Long  ' ������ �����
    
    StepIn
    R = GetRep(SN_LIST)
    MS "��������� � PartnerCenter.Autodesk.com " & R.EOL & " ����� Registered SN"
    
 '---------- ������������� ������� ������ � ���������� ������ -------------
''    iActive = 1: iUpdate = 1
''    nActive = 0: nUpdate = 0: nToClean = 0
''
''    Dim Ractive As TOCmatch, Rupdate As TOCmatch
''
''    Ractive = GetRep(ST_ACTIVE)
''    Rupdate = GetRep(ST_UPDATE)
''    NewSheet Ractive.Name
''    NewSheet Rupdate.Name
    NewSheet SN_UPDATE
    
  '--------------- ���� �������� �� SN_List ---------------
    i = 2
    With Workbooks(R.RepFile).Sheets(R.SheetN)
        Do While i <= R.EOL
            .Activate
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
            SNupdate SN_TMP, SN_UPDATE
            i = i + SNstep
        Loop
    End With
End Sub
Sub SNread(F As String)
'
' - SNread  - ������� ������ ����� �������� ������� �� Output.CSV
'             � ������� ����� ���� ����� F
'   31/1/2012
'   13.08.13 - ������� ��������� - ������ ������ � A_SN_2
'   16.08.13 - ���������� ������������ ����� �� SN

    Dim R As TOCmatch

    NewSheet F      ' ���������� ���� SN �� PartnerCenter
    
    R = GetRep(F)
    If R.EOL <> 1 Then ErrMsg FATAL_ERR, "���� \\W_TMP.A_PC_2' �� ������!"
    Workbooks(R.RepFile).Sheets(R.SheetN).Activate
    
    With ActiveSheet.QueryTables.Add( _
            Connection:="TEXT;" & DownloadDir & "output.csv", _
            Destination:=Cells(2, 1))
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
        On Error GoTo 0
     End With
   
    ActiveSheet.Rows(2).Delete
    ActiveSheet.Columns(SNTMP_ACTDATE_COL).Select
    Selection.NumberFormat = "d/mm/yyyy;@"
    Call SheetSort(R.Name, 1)
    R.EOL = EOL(R.SheetN)
'----- ������������ -----------------
'
' ��������  : =SN :i=R:i-1=Registered
'    NOP    :  0  : - : - :
' Del i-1 Rw:  1  : 0 : 0 : Upgraded -> Upgraded     \ ��������� ������
' Del i-1 Rw:  1  : 0 : 1 : Upgraded -> Registered    |   �� ��������
' Del i Row :  1  : 1 : 0 : Registered -> Upgraded    |   Registered
' Del i-1 Rw:  1  : 1 : 1 : Registered -> Registered /  ��� ���������

    Dim i As Long, SN As String, Status As String
    SN = "": Status = "": i = 2
    With ActiveSheet
        Do
            If .Cells(i, SNTMP_SN_COL) = SN Then
                If .Cells(i, SNTMP_STATUS_COL) <> SNTMP_STATUS_REGISTERED And _
                                         Status = SNTMP_STATUS_REGISTERED _
                Then
                    .Rows(i).Delete
                Else
                    .Rows(i - 1).Delete
                End If
                R.EOL = R.EOL - 1
                i = i - 1
            End If
            SN = .Cells(i, SNTMP_SN_COL)
            Status = .Cells(i, SNTMP_STATUS_COL)
            i = i + 1
        Loop While i <= R.EOL
    End With

    R.Dat = Now
    R.Made = "SNread"
    R.CreateDat = Now
    RepTOC = R: Call WrTOC(R.Name)
    Exit Sub
Rep:
    If MsgBox("� �������� '��������' �� ������ ���� 'output.csv'." _
        & vbCrLf & "'" & DownloadDir & "output.csv" & "'" _
        & vbCrLf & vbCrLf & "�������, ��� �� �������� �� PartnerCenter � �������.") _
        = vbYes Then
            Call SNread(F)
    Else
            End
    End If
End Sub
Sub SNupdate(FrF As String, ToF As String)
'
' - SNupdate(FrF, ToF)   - ������� ������ ������ �� ����� FrF
'                          � ���� ToF � ����� ����� ToF (Append)
' 16.6.13

    Dim Rfr As TOCmatch, Rto As TOCmatch
    
    Rfr = GetRep(FrF)
    Rto = GetRep(ToF)
    
    Workbooks(Rfr.RepFile).Sheets(Rfr.SheetN).Activate
    ActiveSheet.Rows("2:" & Rfr.EOL).Copy _
        Destination:=Workbooks(Rto.RepFile).Sheets(Rto.SheetN).Cells(Rto.EOL + 1, 1)
    
    Rto.Dat = Now
    Rto.Made = "SNupdate"
    Rto.CreateDat = Now
    Rto.EOL = EOL(Rto.SheetN)
    RepTOC = Rto: Call WrTOC(Rto.Name)
End Sub

'''Function SNcheck(ByVal SN As String, SNtmp) As String
''''
'''' - SNcheck(SN, SNtmp) - c����� ������� SN � SN_TMP � SF. ����������
''''       - Active    - SN � ������� � ��� ������ ����������
''''       - InsAct    - ����� SN - �������� � ������ SF
''''       - Update    - �������� ������ � SF, ��������, Upgrade
''''
'''' 15.8.13
'''
'''    Dim i
'''    Dim R As TOCmatch
'''
'''    R = GetRep(SNtmp)
'''
'''    SNcheck = "Active"
'''
'''    With Workbooks(R.RepFile).Sheets(R.SheetN)
'''        For i = 2 To R.EOL
'''            If SN = .Cells(i, 1) Then GoTo SNfound
'''        Next i
'''        SNcheck = "Not Found"
'''        Exit Function
'''SNfound:
'''        Select Case ActiveSNchk(SN, SNtmp)
'''        Case "OK":
'''        Case "ActiveSNupd":
'''        Case Else
'''        End Select
'''
'''        If .Cells(i, 1) = .Cells(i + 1, 1) Then
'''            Stop    ' �� �������� - ������ ���� � SNtmp ��������� ����� �� SN!!!
'''        End If
'''
'''        If .Cells(i, SN_PC_STAT_COL) = "Registered" Then
'''        End If
'''    End With
'''End Function
'''
'''Sub SNsortOut(SN, SNtmp, SNactive, SNupdate)
''''
'''' - SNsortOut(SNtmp, SNactive, SNupdate) - ������ � �������� SN �� SNtmp.
''''      - � SNactive ��������� ������ � Registered SN
''''      - � SNudpade - �� Upgraded SN, ������� ���� �������� � SF
''''   ����� ���� � SN_LIST ���������� ���� SN:
''''           - �����      - SN ������, �� Registered
''''           - ���������� - ������, ��������� Update
''''           - �������� ������ - �� ������ ��� ������� ������ ��������
''''
''''   13.8.2013
'''
'''    Dim iTmp As Long    '- ��������� - ����� ������ � SNtmp
'''
'''    For iTmp = 2 To Rtmp.EOL
'''        SN =
'''    Next iTmp
'''End Sub
'''
'''Sub LoadSNfrOutputCSV()
''''
'''' [3] - ������ SN ���������� �� Licence Inquiry � ������ output.csv
'''' ������ ����� ���� ���������. ����� ��������� ��� ���������.
''''
''''   1.2.2012
'''
'''    Const CSVstamp1 = "Serial Number"
'''    Const CSVstamp2 = "Product Key"
'''    Dim F, SN, Status As String
'''    Dim R, C, i, j As Integer
'''
'''    Start3PASS ("�������� ������ output*.csv")
'''
'''' ������ ������ ������ output*.csv � ������� B ����� 3PASS
'''    F = dir(DownloadDir & "output*.csv")
'''    i = 2
'''    Do While F <> ""
'''        Cells(i, 2) = DownloadDir & F
'''        Cells(i, 2).Select
'''        F = dir()
'''        i = i + 1
'''    Loop
'''
'''' ������ ������ �� ������ � �
'''    R = 5: i = 2: C = 7
'''    Do While Cells(i, 2) <> ""
'''        Cells(i, 2).Select
'''        Call SNread(Cells(i, 2), R, C)    ' ������ ����� output.csv
'''    ' ��������, ���������� �� ���� ���������?
'''        If Cells(R, C) <> CSVstamp1 Or Cells(R, C + 1) <> CSVstamp2 Then
'''            MsgBox "������������ ������� ���� (" & F & _
'''                ") � ����� 3PASS, ������� � ������ (" _
'''                & R & "," & C & ")", vbCritical, "ERROR!"
'''            Stop
'''        End If
'''    ' �������� ����� ������������ �����
'''        Range("G" & R & ":V" & R).Delete Shift:=xlUp
'''        While Cells(R, C) <> "" ' ���������� R �� ������ ������
'''            R = R + 1
'''        Wend
'''        i = i + 1
'''    Loop
'''
'''    ' ���������, ��� �� SN �� ������ � ������� � �������?
'''    R = R - 1
'''    i = 5
'''    Do
'''        SN = Left(Cells(i, 1), 12)        ' ������ SN
''''        Cells(i, 1).Select
'''        For j = 2 To R
'''            If SN = "" Then Exit Do
'''            If SN = Cells(j, 7) Then
'''                i = i + 1
'''                Exit For
'''            End If
'''        Next j
'''    Loop
'''
'''    If i - 1 <> Cells(3, 1) Then
'''        MsgBox "�� ��� �������� ������ ������� -- ��.���� 3PASS", , "ERROR!"
'''        Stop
'''    End If
'''
'''    Range("B5:B" & i).ClearContents
'''    Range("A4:A" & Cells(3, 1)).ClearContents
'''    Range("B4:F" & R).FillDown  ' �������� ������� �� ����� (R)
'''
'''    End3PASS 3
'''End Sub
'''Sub DoDeDupSN()
''''
'''' [*] ������������ �������� ������� �� ����� 3PASS
''''   6.2.2012
'''
'''End Sub
'''Sub WrDL3pass()
''''
'''' [3PASS] - ������ ��� ������ Data Loader'a ����� 3PASS � Salesforce
''''   3/2/2012
'''
'''    Start3PASS "������ ����� 3PASS ����������� DataLoader � SF"
'''
'''    Columns("C:C").Copy
'''    Columns("O:O").PasteSpecial Paste:=xlPasteValues
''''    Columns("O:O").Select
''''    Selection.PasteSpecial Paste:=xlPasteValues
'''
'''    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\ADSK\3PASS\"
'''    WriteCSV A3PASS, "3PASS.txt", 1, 7
'''
'''    Shell "quota_3PASS.bat"
''''    Shell "Copy 3PASS.csv C:\SFconstr"
'''
'''    End3PASS (4)
'''End Sub
