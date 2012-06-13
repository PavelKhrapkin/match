Attribute VB_Name = "HandlIntergraph"
'--------------------------------------создание проектов----------------
' HandlIntergraph - модуль для работы по Intergraph
'
' (*) ConsOppCreate - Создание Проектов SF по данным Match и запись через DL
' (*) PaidConsumers - Запись Платежей по Расходникам через DL
'   14.2.2012

Option Explicit

Public Const NewOppSheet = "O_NewOpp"
Public Const Consumers = "Расходники"
Public Const ConumersActivity = "продажа расходных материалов и ЗИП"
Dim OppN As Integer
Sub ConsOppCreate()
'
' проход по отчету Платежей и создание "bulky" Проектов по расходникам
'   12.2.2012

    Dim i As Integer
    Dim Acc, SalesRep As String
    
    Lines = ModStart(1, "Проекты по Расходникам") - 3
    
    ClearSheet NewOppSheet, 4
    OppN = 1
'    Lines = 250
'------ Расходники Панкова ----------------
    For i = 2 To Lines
        SalesRep = Sheets(1).Cells(i, 22)
        If SalesRep = "Панков" Or SalesRep = "Фролов" Then
            Acc = Sheets(1).Cells(i, 9)
            If IsOpp(Acc, ConumersActivity) = 0 And _
                    Sheets(1).Cells(i, 1) = 1 And _
                    Sheets(1).Cells(i, 4) <> 1 Then
                NewOpp Acc, Acc & "-" & "Расходники", "1.1.2020", "Фролов", _
                    500000, "RUB", "Расходники"
            End If
        End If
    Next i
    
    Columns("E:E").Select           ' подправляем формат Строимости проекта
    Selection.NumberFormat = "0"
    
    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\OppInsert\"
    WriteCSV NewOppSheet, "OppInsert.txt"
    Shell "quota2.bat OppInsert.TXT C:\SFconstr\OppInsert.csv"

    ModEnd NewOppSheet
End Sub


