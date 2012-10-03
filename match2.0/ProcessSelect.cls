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
'   1.10.12 - ��������� ������� ������ "����� ��������"

    Dim i As Long, Proc As String

    If ActiveCell.Value = PROC_START Then
        TraceStop = False
        ProcStart ActiveSheet.Cells(ActiveCell.Row, PROC_NAME_COL)
    ElseIf ActiveCell.Value = "0" Then
'==== ����� ��������
        With ActiveSheet
            i = ActiveCell.Row
            Proc = .Cells(i, PROC_NAME_COL)
            .Range(Cells(i, 1), Cells(i, 3)).Interior.ColorIndex = 0
            If .Cells(i, ActiveCell.Column - 3) = PROC_START Then ProcReset Proc
        End With
    End If
End Sub