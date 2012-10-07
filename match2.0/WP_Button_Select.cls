VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WP_Button_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
'
' H SelectionChange Event Handler
'   7.10.12

    Dim i As Long, Proc As String

    With ActiveCell
        If .Row = 1 Then
            i = .Column
            If i <= .Cells(1, 1) And i >= .Cells(1, 2) Then
                Call xAdapt_Continue(.Value)
            End If
        End If
    End With
End Sub

