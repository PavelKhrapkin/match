VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
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
'   24.10.12 - заменили .Value на V
'   11.11.12 - выход, если DirDBs еще не подготовлен; переписано

    Dim i As Long, j As Long, V As String, Proc As String
    Dim D As String
    
    With ActiveCell
        i = .Column
        j = .Row
    End With
    
    With ActiveSheet
        D = .Cells(1, 5)
        If D = "" Then Exit Sub
        V = .Cells(j, i)
        If j = 1 And i >= .Cells(1, 3) And i <= .Cells(1, 4) Then GoTo x_Continue
        If j >= 25 And i = 11 And V <> "" Then GoTo x_Continue
    End With
    Exit Sub
x_Continue: Application.Run "'" & D, V, j
End Sub
