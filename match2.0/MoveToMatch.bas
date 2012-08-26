Attribute VB_Name = "MoveToMatch"
'---------------------------------------------------------------------------
' �������� ����� MoveToMatch, ������������� � ����� match.xlsm. ��������� ������� ���������.
'
' * MoveInMatch    - ������� �������� ��������� � ���� � ������ Loader'�
'
' �.�.������� 26.8.2012

    Option Explicit    ' Force explicit variable declaration
    
Sub MoveInMatch()
Attribute MoveInMatch.VB_Description = "20.7.12 MoveToMatch of Application match2.0"
Attribute MoveInMatch.VB_ProcData.VB_Invoke_Func = "�\n14"
'
' <*> MoveToMatch() - ����������� �������� ������ � ���� � ������ ��� ���������
'
' �������� ���������� �� Application.Run �� MoveToMatch ������������ �� Ctrl+�
' ������� �������� (���� 1 ��������� �����) ������������ �� ������, ���������������� TOCmatch
'
' 18.8.12 - �������� �� MoveToMatch, ������������� � PERSONAL.xlsb
' 25.8.12 - ������� �������� ����� ���������� � ����� InSheetN ������ ������������ �����
' 26.8.12 - ����� ������� ������ � �������� "1" � ���� Done

    Dim NewRep As String    ' ��� ����� � ����� �������
    Dim i As Long
    Dim IsSF As Boolean     '=TRUE, ���� ������� �������� �� Salesforce
    Dim InSheetN As Integer '���� � TOCmatch- ����� ����� �������� ��������� ��� MoveToMatch
    
    NewRep = ActiveWorkbook.Name
    RepName = ActiveSheet.Name
    Lines = EOL(RepName, Workbooks(NewRep))
    
    GetRep TOC
    
    IsSF = CheckStamp(6, NewRep, Lines)

    For i = 8 To RepTOC.EOL
        InSheetN = 1
        With DB_MATCH.Sheets(TOC)
            If .Cells(i, TOC_INSHEETN) <> "" Then
                InSheetN = .Cells(i, TOC_INSHEETN)
            End If
        End With
        If CheckStamp(i, NewRep, Lines, IsSF, InSheetN) Then GoTo RepNameHandle
    Next i
    FatalRep "MoveToMatch: ���� " & NewRep, RepName
        
'----- ����� ����� ���������. �������� ������� ����� ����� -----
RepNameHandle:
    Dim RepFile As String
    Dim RepLoader As String
    Dim Created As Date
    Dim MyDB As Workbook
    Dim TabColor
    
    With DB_MATCH.Sheets(TOC)
        Lines = Lines - .Cells(i, TOC_RESLINES_COL) '= EOL - �����
        LinesOld = .Cells(i, TOC_EOL_COL)           'EOL ������� ������
        DirDBs = .Cells(1, TOC_F_DIR_COL)
        RepFile = .Cells(i, TOC_REPFILE_COL)
        RepName = .Cells(i, TOC_REPNAME_COL)
        TabColor = .Cells(i, TOC_SHEETN_COL).Interior.Color
    End With
    
    Set MyDB = Workbooks.Open(DirDBs & RepFile, UpdateLinks:=False)
    
    With Workbooks(NewRep).Sheets(InSheetN)
        If RepFile = F_SFDC Then
            Created = Mid(.Cells(Lines + 5, 1), 24)
        ElseIf RepName = PAY_SHEET Or RepName = DOG_SHEET Then
            Created = Right$(.Name, 8)
        ElseIf RepName = Acc1C Then
            Created = Right$(.Cells(1, 1), 8)
        ElseIf RepFile = F_STOCK Then
            Created = MyDB.BuiltinDocumentProperties(12)    '���� ���������� Save
        Else
            Created = "1.1.1900"
        End If
        .UsedRange.Rows.RowHeight = 15
        .Name = "TMP"
        .Move Before:=MyDB.Sheets(RepName)
    End With
    
    With MyDB
        .Activate
        Application.DisplayAlerts = False
        .Sheets(RepName).Delete
        Application.DisplayAlerts = True
        .Sheets("TMP").Name = RepName
        .Sheets(RepName).Tab.Color = TabColor
    End With
    
    LogWr "MoveToMatch: �������� ����� ����� " & RepName _
        & "; EOL=" & Lines & " �����, � ������� " & LinesOld
'------------- match TOC � Log write � Save --------------
    With DB_MATCH.Sheets(TOC)
        .Activate
        .Cells(i, TOC_DATE_COL) = Now
''''''        .Cells(i, TOC_CREATED_COL) = ""
        .Cells(i, TOC_EOL_COL) = Lines
        .Cells(i, TOC_MADE_COL) = REP_LOADED
        RepLoader = .Cells(i, TOC_REPLOADER_COL)
        .Cells(i, TOC_CREATED_COL) = Created
        .Cells(1, 1) = Now
        .Cells(1, TOC_F_DIR_COL) = DirDBs
'----------- ���������� ���� � TOCmatch �� ������� -------------
        Dim d As Date, MaxDays As Integer
        For i = 4 To RepTOC.EOL
            d = .Cells(i, TOC_DATE_COL)
            MaxDays = .Cells(i, TOC_MAXDAYS_COL)
            If d <> "0:00:00" And Now - d > MaxDays Then
                .Cells(i, TOC_DATE_COL).Interior.Color = vbRed
            Else
                .Cells(i, TOC_DATE_COL).Interior.Color = vbWhite
            End If
        Next i
    End With
'---------- ����� ���� ���������, ���������� � ����������� ����������
    With DB_MATCH.Sheets(Process)
        For i = 6 To EOL(Process, DB_MATCH)
            If .Cells(i, PROC_REP1_COL) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 1) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 2) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 3) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 4) = RepName Then
                .Cells(i, PROC_STEPDONE_COL) = ""
                .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 0
            End If
        Next i
    End With
    
    LogWr "����� ����� '" & RepName & "' �������� � " & RepFile
'--- ��������� Loader - ��������� ��������� ������ ������ ---
    If RepLoader <> "" Then
        ProcStart RepLoader
    End If
    MyDB.Save
End Sub
