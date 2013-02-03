Attribute VB_Name = "WPhandling"
'====================================================================
' WPhandling -- ћодуль дл€ работы с листом рабочим листом WP
'   5.11.2012
'
' после выбора данных в Match они обрабатываютс€ на рабочем листе WP

Option Explicit

Sub WP_Paid(Frm As String, Optional InitialPayRow = 2)
'
' S WP_Paid(Frm[,InitialPayRow])    - обработка ѕлатежей по проектам на листе WP
'
'   8.10.12
'   5.11.12 - универсальный характер WP_Paid отражен в названии процедуры;
'           - различныен фильтры дл€ разных Ўаблонов WP

    StepIn
    
    Dim P As TOCmatch
    Dim i As Long
        
    P = GetRep(PAY_SHEET)
    
    With Workbooks(P.RepFile).Sheets(P.SheetN)
        For i = InitialPayRow To P.EOL
            If .Cells(i, PAYINSF_COL) <> 1 Then
                If Frm = "HDR_WP" Then
                    If .Cells(i, PAYISACC_COL) <> "" Then GoTo Go_xAdapt
                Else
                    GoTo Go_xAdapt
                End If
            End If
        Next i
    End With
    
Go_xAdapt:
    With DB_MATCH.Sheets(Process)
        Dim iStep As Long
        iStep = ToStep(.Cells(1, PROCESS_NAME_COL), .Cells(1, STEP_NAME_COL))
        .Cells(iStep, PROC_PAR2_COL) = i + 1
    End With
    xAdapt Frm, i
End Sub
