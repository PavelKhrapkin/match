Attribute VB_Name = "WPhandling"
'====================================================================
' WPhandling -- ������ ��� ������ � ������ ������� ������ WP
'   8.10.2012
'
' ����� ������ ������ � Match ��� �������������� �� ������� ����� WP

Option Explicit

Sub WP_PdOpp(Frm As String, Optional InitialPayRow = 2)
'
' S WP_PaidOpp(Form)    - ��������� �������� �� �������� �� ����� WP
'
'   8.10.12

    StepIn
    
    Dim P As TOCmatch
    Dim i As Long
        
    P = GetRep(PAY_SHEET)
    
    With Workbooks(P.RepFile).Sheets(P.SheetN)
        For i = InitialPayRow To P.EOL
            If .Cells(i, PAYINSF_COL) <> 1 _
                    And .Cells(i, PAYISACC_COL) <> "" Then
                With DB_MATCH.Sheets(WP)
                    .Cells(WP_CONTEXT_LINE, WP_CONTEXT_COL) = i
                    xAdapt Frm
                End With
            End If
        Next i
    End With
End Sub


