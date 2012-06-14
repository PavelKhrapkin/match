Attribute VB_Name = "HandleConsumer"
'--------------------------------------создание проектов----------------
' ConsumerHandling - модуль для работы по Расходникам
'   впоследствии я предполагаю значительно более широкий профиль
'   и даже пеепрофилирование этого модуля
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
Sub NewOpp(Account, OppName, CloseDate, OwnerOpp, Value, CurrencyOpp, TypOpp)
'
' новый проект для записи DL в Организации Account
'   12.2.2012

    Dim i, AccountId, OwId
    
    AccountId = AccId(Account)
    OwId = OwnerId(OwnerOpp)
    Sheets(NewOppSheet).Select
    
    For i = 1 To OppN
        If Cells(i, 1) = AccountId Then Exit Sub    ' дедупликация NewOpp
    Next i
    
    OppN = OppN + 1
    Cells(OppN, 1) = AccountId
    Cells(OppN, 2) = OppName
    Cells(OppN, 3) = CloseDate
    Cells(OppN, 4) = OwId
    Cells(OppN, 5) = Value
    Cells(OppN, 6) = CurrencyOpp
    Cells(OppN, 7) = TypOpp
    If TypOpp = Consumers Then
        Cells(OppN, 8) = "90%-первые деньги пришли на счет"
        Cells(OppN, 9) = ConumersActivity
        Cells(OppN, 10) = "Расходные материалы и ЗИП"
    End If
End Sub


Function Adr1C(Acc) As String
'
'   возвращает адрес Организации Acc в Списке клиентов 1С или ошибку, если ее нет
'   13.2.2012
    On Error Resume Next
    Adr1C = 0
    Adr1C = WorksheetFunction.VLookup(Acc, Sheets(Acc1C).Range("E:H"), 4, False)
    On Error GoTo 0
End Function
Function AccId(Account)
'
' получение Id Организации из отчета SFacc
'   13.2.2012
    On Error Resume Next
    AccId = 0
    AccId = WorksheetFunction.VLookup(Account, Sheets(SFacc).Range("B:C"), 2, False)
    On Error GoTo 0
End Function
Function OwnerId(Owner)
'
' получение Id по фамилии Продавца с переадресацией по таблице "We"
'   13.2.2012
    On Error Resume Next
    OwnerId = 0
    OwnerId = WorksheetFunction.VLookup(Owner, Sheets(We).Range("A8:C21"), 3, False)
    On Error GoTo 0
End Function
Function IsOpp(Account, T)
'
' проверка, есть ли в Организации Account незакрытый проект типа Т.
' Если есть - возвращает номер такого проекта, если нет - 0
'   13.2.2012

    Dim i
    
    IsOpp = 0
    If AccId(Account) = 0 Then Exit Function    ' если Асс нет в SF -> =0
    For i = 2 To EOL(SFopp)
        If Cells(i, 4) = Account And Cells(i, 6) = 0 And Cells(i, 17) = T Then
            IsOpp = Cells(i, 2)                 ' Есть Проект нужного типа и он не закрыт
            Exit For
        End If
    Next i
End Function

Sub PaidConsumers()
'
' проход по списку Платежей и обработка Расходников
'   14.2.2012

    Dim i, j, a, Pct
    Dim Account As String
    Const SheetN = "P_PaidConsumer"
    
    ClearSheet SheetN, 7
    Lines = ModStart(1, "Обработка Платежей по расходникам")
    
    j = 2
    For i = 2 To Lines
        Call Progress(i / Lines)
'        Sheets(1).Select
'        Cells(i, 1).Select
        Account = Sheets(1).Cells(i, 9)
            ' есть Проект и Платеж не пустой и его нет в SF
        If IsOpp(Account, ConumersActivity) <> 0 And _
            Sheets(1).Cells(i, 4) <> 1 And _
            Trim(Sheets(1).Cells(i, 6)) <> "" Then
                Sheets(SheetN).Cells(j, 1) = Sheets(1).Cells(i, 6)              ' Плат.док.
                Sheets(SheetN).Cells(j, 2) = DDMMYYYY(Sheets(1).Cells(i, 7))    ' Дата
                Sheets(SheetN).Cells(j, 3) = Sheets(1).Cells(i, 8)              ' Счет
                Sheets(SheetN).Cells(j, 4) = Dec(Sheets(1).Cells(i, 18))        ' Итог руб.
                Sheets(SheetN).Cells(j, 5) = Sheets(1).Cells(i, 19)             ' Товар
                Sheets(SheetN).Cells(j, 6) = ContractId( _
                    Sheets(1).Cells(i, 25), Sheets(1).Cells(i, 26))             ' ContractId
                Sheets(SheetN).Cells(j, 7) = IsOpp(Account, ConumersActivity)   ' OppN
             Sheets(SheetN).Select
             Cells(j, 1).Select
             j = j + 1
        End If
    Next i
    
' !!!! проверять, нет ли новых платежей среди "красных" в SF
'    Lines = EOL(SFD)
'    For i = 2 To EOL(SheetN)
'
    
    ChDir "C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\Payment\"
    WriteCSV SheetN, "Paid.txt"
    Shell "quota2.bat"

    ModEnd 1
End Sub
Function DDMMYYYY(d) As String
'
' преобразование даты d в текстовый формат DDMMYYYY
'   14.2.2012
    DDMMYYYY = Day(d) & "." & Month(d) & "." & Year(d)
End Function
Function Dec(a) As String
'
' преобразование числа а в текстовый формат с десятичной точкой
'   14.2.2012

    Dec = "'" & WorksheetFunction.Substitute(a, ",", ".")
'    On Error Resume Next
'     If WorksheetFunction.Search(",", Dec) <> "#VALUE!" Then MsgBox "запятая!"
'    On Error GoTo 0
End Function
Function ContractId(Contr, Optional MainContr = "") As String
'
' получение Id по имени контракта и основному контракту, если есть
' если контракта нет в SF, возвращает 0
'   14.2.2012

    Dim C As String
    
    C = MainContr
    If C = "" Then
        C = Contr
    Else
        C = C & "/" & Contr
    End If
    On Error Resume Next
    ContractId = ""
    ContractId = WorksheetFunction.VLookup(C, Sheets(SFD).Range("2:15"), 13, False)
    On Error GoTo 0
End Function
