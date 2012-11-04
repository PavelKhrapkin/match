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
'   1.10.12 - обработка голубой кнопки "Сброс процесса"
'   5.11.12 - инициализация TraseStep и TraceStop

    Dim i As Long, Proc As String

    If ActiveCell.Value = PROC_START Then
        TraceStop = False: TraceStep = False
        ProcStart ActiveSheet.Cells(ActiveCell.Row, PROC_NAME_COL)
    ElseIf ActiveCell.Value = "0" Then
'==== Сброс Процесса
        With ActiveSheet
            i = ActiveCell.Row
            Proc = .Cells(i, PROC_NAME_COL)
            .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 0
            If .Cells(i, ActiveCell.Column - 3) = PROC_START Then
                TraceStop = False: TraceStep = False
                ProcReset Proc
            End If
        End With
    End If
End Sub
