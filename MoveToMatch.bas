Attribute VB_Name = "MoveToMatch"
'---------------------------------------------------------------------------
' �������� ����� MoveToMatch, ������������� � ����� match.xlsm. ��������� ������� ���������.
'
' * MoveInMatch    - ������� �������� ��������� � ���� � ������ Loader'�
'
' �.�.������� 27.8.2013

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
'  6.4.13 - ����� ��� ������� ��������� � match ���� �� ������ ���� ������
' 13.5.13 - ���������� ������-����������� � TOCmatch
' 17.8.13 - �������� ������� � ��������� ���������� ���
' 18.8.13 - ResLines ������ ����� ��� 2 / 7
' 23.8.13 - SheetSort ������������ ���������, ���� ��� ����� �������
' 24.8.13 - ��������� InSheetN � TOC. ������ �������� ������ ������ ���� � ����� 1
' 27.8.13 - ������������ ������������� ���������� ��������� RepTOC
    
    Dim NewRep As String    ' ��� ����� � ����� �������
    Dim i As Long
    Dim IsSF As Boolean     '=TRUE, ���� ������� �������� �� Salesforce
    Dim IsPartialUpdate     '=TRUE, ���� ������� �������� �������� ���� ����� ������
    Dim FrDateTOC As Date, ToDateTOC As Date, NewFrDate As Date, NewToDate As Date
    Dim NewFrDate_Row As Long, NewFrDate_Col As Long
    Dim NewToDate_Row As Long, NewToDate_Col As Long
    Dim InSheetN As Integer '���� � TOCmatch- ����� ����� �������� ��������� ��� MoveToMatch
    Dim LocalTOC As TOCmatch
    
    IsPartialUpdate = False
    NewRep = ActiveWorkbook.Name
    RepName = ActiveSheet.Name
    Lines = EOL(RepName, Workbooks(NewRep))
    
    LocalTOC = GetRep(TOC)
    
    IsSF = CheckStamp(6, NewRep, Lines)

    With DB_MATCH.Sheets(TOC)
        For i = TOCrepLines To LocalTOC.EOL
            If .Cells(i, TOC_REPNAME_COL) = "" Then GoTo NxDoc
            InSheetN = 1
''            If .Cells(i, TOC_INSHEETN) <> "" Then
''                InSheetN = .Cells(i, TOC_INSHEETN)
''            End If
            If CheckStamp(i, NewRep, Lines, IsSF, InSheetN) Then GoTo RepNameHandle
NxDoc:  Next i
    End With
    GoTo FatalInFile
        
'----- ����� ����� ���������. �������� ������� ����� ����� -----
RepNameHandle:
    Dim RepFile As String
    Dim RepLoader As String
    Dim Created As Date
    Dim MyDB As Workbook
    Dim TabColor
    
    With DB_MATCH.Sheets(TOC)
    
        If NewRep = .Cells(i, TOC_REPFILE_COL) Then
            MS "��� ���� ���� ������ match! ��� �� ���� ���������."
            End
        End If
        RepName = .Cells(i, TOC_REPNAME_COL)
        RepFile = .Cells(i, TOC_REPFILE_COL)
         'Lines = EOL - �����
        Lines = Lines - GetReslines(RepName, True, .Cells(i, TOC_RESLINES_COL))
        LinesOld = .Cells(i, TOC_EOL_COL)           'EOL ������� ������
        DirDBs = .Cells(1, TOC_F_DIR_COL)
        TabColor = .Cells(i, TOC_SHEETN_COL).Interior.Color
      '--��������� ��������� ��� � match � ����� ������ ---
        FrDateTOC = .Cells(i, TOC_FRDATE_COL)   ' ���� �������� ������
        ToDateTOC = .Cells(i, TOC_TODATE_COL)   '.. � Match
        NewFrDate_Row = .Cells(i, TOC_FRDATEROW_COL)
        NewFrDate_Col = .Cells(i, TOC_DATECOL_COL)
        Dim ToStr As String
        ToStr = .Cells(i, TOC_TODATEROW_COL)
        If ToStr = "EOL" Then
            NewToDate_Row = Lines
        ElseIf WorksheetFunction.IsNumber(ToStr) Then
            NewToDate_Row = ToStr
        End If
        NewToDate_Col = .Cells(i, TOC_DATECOL_COL)
        
    End With
    
    Set MyDB = Workbooks.Open(DirDBs & RepFile, UpdateLinks:=False)
    
    With Workbooks(NewRep).Sheets(InSheetN)
        If RepFile = F_SFDC Then
            Dim tst As String
            tst = .Cells(Lines + 5, 1)
            Created = GetDate(Right(.Cells(Lines + 5, 1), 16))
        ElseIf RepName = PAY_SHEET Or RepName = DOG_SHEET Then
            .Activate
            .Rows("1:" & Lines).AutoFilter
            DateCol InSheetN, NewToDate_Col
            SheetSort InSheetN, NewToDate_Col
            Created = GetDate(Right$(.Name, 8))
            Dim DateCell As String
            Do
                DateCell = .Cells(NewFrDate_Row, NewFrDate_Col)
                If IsDate(DateCell) Then
                    Exit Do
                Else
                    NewFrDate_Row = NewFrDate_Row + 1
                    If NewFrDate_Row > Lines Then GoTo FatalFrDate
                End If
            Loop
            NewFrDate = GetDate(DateCell)
            Do
                DateCell = .Cells(NewToDate_Row, NewToDate_Col)
                If IsDate(DateCell) Then
                    Exit Do
                Else
                    NewToDate_Row = NewToDate_Row - 1
                    If NewToDate_Row < NewFrDate_Row Then GoTo FatalToDate
                End If
            Loop
            NewToDate = GetDate(DateCell)
            If NewFrDate > NewToDate Then GoTo FatalFrToDate
            If NewFrDate <> FrDateTOC Or NewToDate < ToDateTOC Then
                IsPartialUpdate = True
            End If
        ElseIf RepName = Acc1C Then
            Created = GetDate(Right$(.Cells(1, 1), 8))
        ElseIf RepFile = F_STOCK Then
            Created = GetDate(MyDB.BuiltinDocumentProperties(12))   '���� ���������� Save
        Else
            Created = "0:0"
            NewFrDate = "0:0": NewToDate = "0:0"
        End If
        .UsedRange.Rows.RowHeight = 15
        .Name = "TMP"
        .Move Before:=MyDB.Sheets(RepName)
    End With
    
    With MyDB
        .Activate
  '-- ���� ��������� ���������� - ������� ����� �� �������, � ���������������
  '-- .. ��� � *_OLD, ����� ����� ����� �� � ���� MergeRep Loader'�.
  '-- .. ���� _OLD ��� ����, �� ��� �� ��������� - ���������� ������� "���������" �����
        If IsPartialUpdate Then
            Dim OldRepName As String, sht As Worksheet
            OldRepName = RepName & "_OLD"
            If SheetExists(OldRepName) Then GoTo DelRep
            .Sheets(RepName).Name = OldRepName
        End If
DelRep: If SheetExists(RepName) Then
            Application.DisplayAlerts = False
            .Sheets(RepName).Delete
            Application.DisplayAlerts = True
        End If
        .Sheets("TMP").Name = RepName
        .Sheets(RepName).Tab.Color = TabColor
    End With
    
'------------- match TOC � Log write � Save --------------
    With DB_MATCH.Sheets(TOC)
        .Activate
        .Cells(i, TOC_DATE_COL) = Now
        .Cells(i, TOC_EOL_COL) = Lines
        .Cells(i, TOC_MADE_COL) = REP_LOADED
        RepLoader = .Cells(i, TOC_REPLOADER_COL)
        .Cells(i, TOC_CREATED_COL) = Created
        If NewFrDate_Col > 0 Then
            .Cells(i, TOC_NEW_FRDATE_COL) = NewFrDate
            .Cells(i, TOC_NEW_TODATE_COL) = NewToDate
        End If
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
    Dim PartStatus As String
    PartStatus = vbCrLf & "��� "
    If IsPartialUpdate Then
        PartStatus = PartStatus & "���������� ����� ������."
    Else
        PartStatus = PartStatus & "������ ��������."
    End If
    LogWr "MoveToMatch: � ���� '" & RepFile & "' �������� ����� ����� '" _
        & RepName & "'; EOL=" & Lines & " �����, � ������� " & LinesOld _
        & PartStatus
        
'--- ��������� Loader - ��������� ��������� ������ ������ ---
    If RepLoader <> "" Then
        ProcStart RepLoader
    End If
    MyDB.Save
    Exit Sub
    Dim Msg As String
FatalInFile:    Msg = "�� ������ �����": GoTo FatMsg
FatalFrDate:    Msg = "FrDate": GoTo FatErMsg
FatalToDate:    Msg = "ToDate"
FatErMsg:       Msg = " �� ���� � ������ " & Msg & "='" & DateCell & "'": GoTo FatMsg
FatalFrToDate:  Msg = " �������� ���� �������� ��������� '" & NewRep _
                    & "': NewFrDate=" & NewFrDate & " < " & "NewToDate=" & NewToDate
FatMsg: ErrMsg FATAL_ERR, "MoveToMatch: " & Msg & vbCrLf & "������� �������� " & NewRep
End Sub
Sub StepReset(iStep)
'
' - StepReset(iStep) - ����� ���� � ������� ��������� - ��������!
' 28.8.12
'  9.9.12 - bug fix � ����� ������������ ���� ��� �������� ������ ���������
' 13.9.12 - bug fix - �� ���������� ���� <*>ProcStart

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
                If .Cells(i, PROC_STEPDONE_COL) = "1" Then ' ���������� <*>ProcStart
                    .Cells(i, PROC_STEPDONE_COL) = ""
                End If
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
