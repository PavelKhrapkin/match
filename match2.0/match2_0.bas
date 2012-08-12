Attribute VB_Name = "match2_0"
'---------------------------------------------------------------------------
' ������� ��� ������ � ������ ������� �� 1� � Salesforce Match SF-1C.xlms
'
' * MoveToMatch    - ������� ����� �� ������ ����� Match1SF    (Ctrl/Shift/M)
' * TriggerOptionsFormulaStyle  - ������������ ���� A1/R1C1    (Ctrl/Shift/R)
'
' �.�.������� 2.1.2012
'   28.1.2012 - ������ �� �������������� ���� � ������� ������
'    5.2.2012 - � MoveToMatch - ������������� �������� ������ �� ������
'   16.5.2012 - �������� ����� SF_PA
'    2.6.2012 - TriggerOptionsFormulaStyle A1/R1C1
'   26.7.2012 - match 2.0 - MoveToMatch � �������������� TOCmatch
'   11.8.2012 - ��������� ��������� - Loader'�� � ProcessEngine

    Option Explicit    ' Force explicit variable declaration
    
Sub MoveToMatch()
Attribute MoveToMatch.VB_Description = "20.7.12 MoveToMatch of Application match2.0"
Attribute MoveToMatch.VB_ProcData.VB_Invoke_Func = "�\n14"
'
' <*> MoveToMatch() - ����������� �������� ������ � ���� � ������ ��� ���������
'
' Keyboard Shortcut: Ctrl+�     -- Ctrl/� �����������, ����� �� ������������ Shift,
'                                  ��������������� ���������� �� Open
'
'Pavel Khrapkin 23-Dec-2011
' 8.2.2012 - ���������� ����� �����, ��������� ��� ���������
' 26.7.12 - match2.0 - ������������� ������ �� ���
' 1.8.12 - RepTOC.EOL ������ ������ EOL(TOC,DB_MATCH), bug fix
'          ����� ���� ���������, ���������� � ����������� ����������
' 11.8.12 - bug fix - ��������� ���� ������

    Dim NewRep As String    ' ��� ����� � ����� �������
    Dim i As Integer
    
    NewRep = ActiveWorkbook.Name
    Lines = EOL(ActiveSheet.Name, Workbooks(NewRep))
    
    GetRep TOC
    
    For i = 4 To RepTOC.EOL
        If IsThisStamp(i, NewRep) Then GoTo RepNameHandle
    Next i
    GoTo FatalNewRep
        
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
    
    With Workbooks(NewRep).Sheets(1)
        If RepFile = F_SFDC Then
            Created = Mid(.Cells(Lines + 5, 1), 24)
        ElseIf RepName = PAY_SHEET Or RepName = DOG_SHEET Then
            Created = Right$(.Name, 8)
        ElseIf RepName = Acc1C Then
            Created = Right$(.Cells(1, 1), 8)
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
        .Cells(i, TOC_CREATED_COL) = ""
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
        For i = 6 To EOL(Process, DB_MATCH)
            If .Cells(i, PROC_REP1_COL) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 1) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 2) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 3) = RepName _
                    Or .Cells(i, PROC_REP1_COL + 4) = RepName Then
                .Cells(i, PROC_STEPDONE_COL) = ""
            End If
        Next i
    End With
    
    LogWr "����� ����� '" & RepName & "' �������� � " & RepFile
    DB_MATCH.Save
    DB_MATCH.Close
'--- ��������� Loader - ��������� ��������� ������ ������ ---
    If RepLoader <> "" Then
        Application.Run "'" & DirDBs & F_MATCH & "'!ProcStart", RepLoader
    End If
    MyDB.Save
    Close
    End
FatalNewRep:
    ErrMsg FATAL_ERR, "������� ����� '" & NewRep & "' �� ���������"
End Sub
Function IsThisStamp(iTOC, NewRep) As Boolean
'
' - IsThisStamp(iTOC) - �������� ������������ ������ ������ ������ � ������ iTOC.
'                       ���� ����� �� ��������� - ����������� �������� � iTOC + 1
' 23.7.2012

    Dim NewRepStamp As String       ' ����� ������ ������
    Dim RepFile As String
    Dim Stamp As String         '= ������ - �����
    Dim StampType As String     '��� ������: ������ (=) ��� ���������
    Dim Stamp_R As Integer      '����� ������, ��� �����
    Dim Stamp_C As Integer      '����� �������, ��� �����
    
    IsThisStamp = False
            
    With DB_MATCH.Sheets(TOC)
        Stamp = .Cells(iTOC, TOC_STAMP_COL)
        RepFile = .Cells(iTOC, TOC_REPFILE_COL)
        If Stamp = "" Then Exit Function        ' ����������� ����� - �� �������!
        StampType = .Cells(iTOC, TOC_STAMP_TYPE_COL)
        Stamp_R = .Cells(iTOC, TOC_STAMP_R_COL)
        If RepFile = F_SFDC Then        ' ������ � ������� SFDC ����� � �����
            Stamp_R = Stamp_R + Lines - .Cells(iTOC, TOC_RESLINES_COL)
        End If
        Stamp_C = .Cells(iTOC, TOC_STAMP_C_COL)
        NewRepStamp = Workbooks(NewRep).Sheets(1).Cells(Stamp_R, Stamp_C)

        If StampType = "=" Then
            If NewRepStamp <> Stamp Then Exit Function
        ElseIf StampType = "I" Then
            If InStr(LCase$(NewRepStamp), LCase$(Stamp)) = 0 Then Exit Function
        Else:
            ErrMsg FATAL_ERR, "���� � ��������� TOCmatch: ��� ������ =" & StampType
        End If

        If .Cells(iTOC, TOC_PARCHECK_COL) <> "" Then    ' ���� ParCheck �� ������ -
            IsThisStamp = IsThisStamp(iTOC + 1, NewRep) ' .. ����������� ��������
        End If
    End With
    
    If RepFile = F_SFDC Then
        IsThisStamp = IsThisStamp(5, NewRep)   '���.�������� ����� ������� SFDC
    End If

    IsThisStamp = True

End Function
Sub TriggerOptionsFormulaStyle()
Attribute TriggerOptionsFormulaStyle.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' * Trigger Options-Formula Style A1/R1C1
'
' CTRL+Shift+R
'
' 2.6.12
    If Application.ReferenceStyle = xlR1C1 Then
        Application.ReferenceStyle = xlA1
    Else
        Application.ReferenceStyle = xlR1C1
    End If
End Sub
