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
'----------------------------------------------------------------
' WP_Button_Select.cls  ЗАПИСЫВАЕТСЯ в WP_PROTOTYPE
'
' H  SelectionChange Event Handler
'   11.10.12
'   20.10.12 - добавлена обработка кнопок Шаблона Select
'   22.10.12 - DirDBs сдвинуто в ячейку [1,5]

    Dim i As Long, Proc As String

    With ActiveCell
        i = .Column
        If .Row = 1 Then
            With ActiveSheet
                If i >= .Cells(1, 3) And i <= .Cells(1, 4) Then
                    Application.Run "'" & ActiveSheet.Cells(1, 5), .Value, .Row
                End If
            End With
        ElseIf .Row >= 25 And i = 10 Then
            Application.Run "'" & ActiveSheet.Cells(1, 5), .Value, .Row
        End If
    End With
End Sub
