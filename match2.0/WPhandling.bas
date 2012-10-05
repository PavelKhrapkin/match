Attribute VB_Name = "WPhandling"
'====================================================================
' WPhandling -- Модуль для работы с листом рабочим листом WP
'   4.10.2012
'
' после выбора данных в Match они обрабатываются на рабочем листе WP

Option Explicit

Sub WP_PdOpp(Frm As String)
'
' S WP_PaidOpp(Form)    - обработка Платежей по проектам на листе WP
'
'   4.10.12

    Const WP_PAY_LINE = 5, WP_GREY_COL = 2

    StepIn
    
    Dim P As TOCmatch
    Dim i As Long
        
    P = GetRep(PAY_SHEET)
    
    With Workbooks(P.RepFile).Sheets(P.SheetN)
        For i = 2 To P.EOL
            If .Cells(i, PAYINSF_COL) <> 1 _
                    And .Cells(i, PAYISACC_COL) <> "" Then
                With DB_MATCH.Sheets(WP)
                    .Cells(WP_PAY_LINE, WP_GREY_COL) = i
                    xAdapt Frm
                End With
            End If
        Next i
    End With
End Sub


