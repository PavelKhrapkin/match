Attribute VB_Name = "WPhandling"
'====================================================================
' WPhandling -- Модуль для работы с листом рабочим листом WP
'   8.10.2012
'
' после выбора данных в Match они обрабатываются на рабочем листе WP

Option Explicit

Sub WP_PdOpp(Frm As String, Optional InitialPayRow = 2)
'
' S WP_PdOpp(Frm[,InitialPayRow])    - обработка Платежей по проектам на листе WP
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
                xAdapt Frm, i
            End If
        Next i
    End With
End Sub


