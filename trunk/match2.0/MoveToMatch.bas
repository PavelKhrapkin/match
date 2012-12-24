Attribute VB_Name = "MoveToMatch"
'---------------------------------------------------------------------------
' �������� ����� MoveToMatch, ������������� � ����� match.xlsm. ��������� ������� ���������.
'
' * MoveInMatch    - ������� �������� ��������� � ���� � ������ Loader'�
'
' �.�.������� 22.12.2012

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
' 28.8.12 - ����� �����, ��������� � �������������� ������������ ���������
' 20.9.12 - Created Date -- ���������� ��� ������� SF
' 22.12.12 - Created Date - ������� �������� � ��. ������ � �������
    
    Dim NewRep As String    ' ��� ����� � ����� �������
    Dim i As Long
    Dim IsSF As Boolean     '=TRUE, ���� ������� �������� �� Salesforce
    Dim InSheetN As Integer '���� � TOCmatch- ����� ����� �������� ��������� ��� MoveToMatch
    
    NewRep = ActiveWorkbook.Name
    RepName = ActiveSheet.Name
    Lines = EOL(RepName, Workbooks(NewRep))
    
    GetRep TOC
    
    IsSF = CheckStamp(6, NewRep, Lines)

    For i = TOCrepLines To RepTOC.EOL
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
            Dim tst As String
            tst = .Cells(Lines + 5, 1)
            Created = GetDate(Right(.Cells(Lines + 5, 1), 16))
        ElseIf RepName = PAY_SHEET Or RepName = DOG_SHEET Then
            Created = GetDate(Right$(.Name, 8))
        ElseIf RepName = Acc1C Then
            Created = GetDate(Right$(.Cells(1, 1), 8))
        ElseIf RepFile = F_STOCK Then
            Created = GetDate(MyDB.BuiltinDocumentProperties(12))   '���� ���������� Save
        Else
            Created = "0:0"
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
        Dim D As Date, MaxDays As Integer
        For i = 4 To RepTOC.EOL
            D = .Cells(i, TOC_DATE_COL)
            MaxDays = .Cells(i, TOC_MAXDAYS_COL)
            If D <> "0:00:00" And Now - D > MaxDays Then
                .Cells(i, TOC_DATE_COL).Interior.Color = vbRed
            Else
                .Cells(i, TOC_DATE_COL).Interior.Color = vbWhite
            End If
        Next i
    End With
'---------- ����� ���� ���������, ���������� � ����������� ����������
    With DB_MATCH.Sheets(Process)
        .Activate
        For i = 6 To EOL(Process, DB_MATCH)
            If .Cells(i, PROC_REP1_COL) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 1) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 2) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 3) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 4) = RepName Then
                StepReset i
            End If
        Next i
    End With
    LogWr "MoveToMatch: � ���� '" & RepFile & "' �������� ����� ����� '" _
        & RepName & "'; EOL=" & Lines & " �����, � ������� " & LinesOld
'--- ��������� Loader - ��������� ��������� ������ ������ ---
    If RepLoader <> "" Then
        ProcStart RepLoader
    End If
    MyDB.Save
End Sub
Sub StepReset(iStep)
'
' - StepReset(iStep) - ����� ���� � ������� ��������� - ��������!
' 28.8.12
'  9.9.12 - bug fix � ����� ������������ ���� ��� �������� ������ ���������

    Dim Step As String, PrevStep As String
    Dim Proc As String, ThisProc As String
    Dim i As Integer, iProc As Integer
    
    With DB_MATCH.Sheets(Process)
        If .Cells(iStep, PROC_STEPDONE_COL) = "" Then Exit Sub
        Step = .Cells(iStep, PROC_STEP_COL)
'---- ����� ���� iStep � ������� ������ ��� ��������� "<*>ProcStart"
        For i = 6 To EOL(Process, DB_MATCH)
            If .Cells(i, PROC_STEP_COL) = PROC_START Then iProc = i
            If i = iStep Then
                .Cells(i, PROC_STEPDONE_COL) = ""
                .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 0
                .Range(Cells(iProc, 1), Cells(iProc, 3)).Interior.ColorIndex = 0
                Exit For
            End If
        Next i
'---- ����� ������� ����� ��������� "<*>ProcEnd"
        For i = iProc + 1 To EOL(Process, DB_MATCH)
            If .Cells(i, PROC_STEP_COL) = PROC_END Then
                .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 0
                Exit For
            End If
        Next i
'---- ����� �����, � ������� � PrevStep ��������� �� ��� � "�����" ��������
        Proc = .Cells(iProc, PROC_NAME_COL)
        For i = iProc + 1 To EOL(Process, DB_MATCH)
            PrevStep = .Cells(i, PROC_PREVSTEP_COL)
            If InStr(PrevStep, Step) <> 0 And i <> iStep Then
                StepReset i                                     '* �������� *
            End If
            If .Cells(i, PROC_STEP_COL) = PROC_END Then Exit For
        Next i
'---- ����� �����, � ������� � PrevStep ��������� �� ��� �� "�������" ��������
        For i = 2 To EOL(Process, DB_MATCH)
            PrevStep = .Cells(iStep, PROC_PREVSTEP_COL)
            ThisProc = .Cells(iStep, PROC_NAME_COL)
            If InStr(PrevStep, Proc & "/" & Step) Then StepReset i '* �������� *
        Next i
    End With
End Sub
