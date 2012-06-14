Attribute VB_Name = "HandlIntergraph"
'--------------------------------------�������� ��������----------------
' HandlIntergraph - ������ ��� ������ �� Intergraph
'
' (*) ConsOppCreate - �������� �������� SF �� ������ Match � ������ ����� DL
' (*) PaidConsumers - ������ �������� �� ����������� ����� DL
'   14.2.2012

Option Explicit

Public Const NewOppSheet = "O_NewOpp"
Public Const Consumers = "����������"
Public Const ConumersActivity = "������� ��������� ���������� � ���"
Dim OppN As Integer
Sub ConsOppCreate()
'
' ������ �� ������ �������� � �������� "bulky" �������� �� �����������
'   12.2.2012

    Dim i As Integer
    Dim Acc, SalesRep As String
    
    Lines = ModStart(1, "������� �� �����������") - 3
    
    ClearSheet NewOppSheet, 4
    OppN = 1
'    Lines = 250
'------ ���������� ������� ----------------
    For i = 2 To Lines
        SalesRep = Sheets(1).Cells(i, 22)
        If SalesRep = "������" Or SalesRep = "������" Then
            Acc = Sheets(1).Cells(i, 9)
            If IsOpp(Acc, ConumersActivity) = 0 And _
                    Sheets(1).Cells(i, 1) = 1 And _
                    Sheets(1).Cells(i, 4) <> 1 Then
                NewOpp Acc, Acc & "-" & "����������", "1.1.2020", "������", _
                    500000, "RUB", "����������"
            End If
        End If
    Next i
    
    Columns("E:E").Select           ' ����������� ������ ���������� �������
    Selection.NumberFormat = "0"
    
    ChDir "C:\Users\������������\Desktop\������ � Match\SFconstrTMP\OppInsert\"
    WriteCSV NewOppSheet, "OppInsert.txt"
    Shell "quota2.bat OppInsert.TXT C:\SFconstr\OppInsert.csv"

    ModEnd NewOppSheet
End Sub


