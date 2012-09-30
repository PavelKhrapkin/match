VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
'
' H SelectionChange Event Handler
'   25/9/12
'   30.9.12 - обработка голубой кнопки "Сброс процесса"

    Dim i As Long

    If ActiveCell.Value = PROC_START Then
        TraceStop = False
        ProcStart ActiveSheet.Cells(ActiveCell.Row, PROC_NAME_COL)
    ElseIf ActiveCell.Value = "0" Then
'==== Сброс Процесса
        With ActiveSheet
            i = ActiveCell.Row
            .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 0
            If .Cells(i, ActiveCell.Column - 3) = PROC_START Then ProcReset i
        End With
    End If
End Sub
