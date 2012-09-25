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
'       25/9/12

    If ActiveCell.Value = PROC_START Then
        TraceStop = False
        ProcStart ActiveSheet.Cells(ActiveCell.Row, PROC_NAME_COL)
    End If
End Sub
