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
' WP_Button_Select.cls  «јѕ»—џ¬ј≈“—я в WP_PROTOTYPE
'
' H  SelectionChange Event Handler
'   11.10.12
'   20.10.12 - добавлена обработка кнопок Ўаблона Select
'   22.10.12 - DirDBs сдвинуто в €чейку [1,5]
'   24.10.12 - заменили .Value на V
'   11.11.12 - выход, если DirDBs еще не подготовлен; переписано
'   17.11.12 - проверка, что комментарий кнопки начинаетс€ с "Button"

    Dim i As Long, j As Long, V As String
    Dim D As String
    Const BUTTON = "Button"
    
    With ActiveCell
        i = .Column
        j = .Row
    End With
    
    With ActiveSheet
        ' считаем, что описатель в заголовке листа
        ' и сама кнопка должны быть непусты
        D = .Cells(1, 5)
        If D = "" Then Exit Sub
        V = .Cells(j, i)
        If V = "" Then Exit Sub
        
        On Error GoTo NoButton
        If Left(.Cells(j, i).Comment.Text, Len(BUTTON)) = BUTTON Then
            On Error GoTo 0
            GoTo x_Continue
        End If
NoButton:
        On Error GoTo 0
        
'        If .Cells(j, i).Comment.Text <> "кнопка" Then Exit Sub
'        If j = 1 And i >= .Cells(1, 3) And i <= .Cells(1, 4) Then GoTo x_Continue
'        If j >= 25 And i = 11 And V <> "" Then GoTo x_Continue
    End With
    GoTo endSub
x_Continue: Application.Run "'" & D, V, j
endSub:
End Sub
