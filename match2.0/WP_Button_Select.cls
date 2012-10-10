VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
'
' H SelectionChange Event Handler
'   11.10.12

    Dim i As Long, Proc As String

    With ActiveCell
        If .Row = 1 Then
            i = .Column
            With ActiveSheet
                If i >= .Cells(1, 3) And i <= .Cells(1, 4) Then
                    Application.Run "'" & .Cells(1, 1), ActiveCell.Value
                End If
            End With
        End If
    End With
End Sub
