Attribute VB_Name = "SFanalitics"
'--------------------------------------------------------------------
' SFanalitics  - Подпрограммы аналитики объектов Salesforce.com:
'       Организаций = Accounts      = Acc   Отчет SFacc
'       Проектов    = Opportunity   = Opp   Отчет SFopp или SF
'       Платежей -- для занесенных в SF --  Отчет SF
'       Договоров   = Contract      = Contr Отчет SFD
'   29.12.2012
'
' - AccId(Account)              - Id SF Организации по имени 1С
' - OwnerId(Owner, Buddy)       - Id SF владельца с переадресацией по таблице We
' - OppByPay(PayKod)            - получает Имя Проекта SF по Платежу 1С
' - OppNbyPay(PayKod)           - получает номер Проекта SF по Платежу 1С
' - OppNbyId(OppId)             - получение номера Проекта OppN по его Id в SF
' - OppIdbyPay(PayKod)          - получает Id SF Проекта по Платежу 1С
' - OppOwner(PayKod)            - возвращает владельца Проекта SF по коду Платежа 1С
' - OppNameById(OppId)          - получение имени Проекта по его Id в SPopp
'??' - IsOpp(Sale, Account, T)     - проверяет, есть ли незавершенный Проект типа Т
' - ContrNbyPay(PayKod)         - получение номера Договора по коду Платежа из диста SF
' - ContractId(ContrCod)        - получение Id SF по контрактe в 1С
' - ContrCod(Contr, MainContr)  - формирование строки <Осн.договор/Договор>
' - ContrCodByOppId(OppId)      - возвращает код Договора (Осн/Договор) по Id Проекта в SF
' - ContrOK(OppN, ContrK)       - True если в SF Платеж связан с Проектом и Договором как в 1С
' - PayIdByK(PayK)  - получение Id SF по коду Платежа

Option Explicit

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
' получение Id SF Организации из отчета SFacc по имени в 1С
'   13.2.2012
    On Error Resume Next
    AccId = 0
    AccId = WorksheetFunction.VLookup(Account, _
        Sheets(SFacc).Range("B:C"), 2, False)
    On Error GoTo 0
End Function
Function OwnerId(Owner, Buddy) As String
'
' получение SF OwnerId по фамилии Продавца в Платежах и Договорах 1C
'   3.3.2012
'   6.3.12 - переход к фамилии Продавца Owner в We
'   12.3.12 - определение Компаньона OppBuddy
    
    Dim X As Range

    For Each X In Range("Продавцы").Rows
        If InStr(Owner, X.Cells(1, 1)) <> 0 Then GoTo Found
    Next X
    LogWr "ERROR! Не найден Продавец " & Owner & _
        ". Это ошибка в Match/We или новый сотрудник в 1С"
    Stop
Found:
'    If InStr(Owner, "Тучков") Then
'''        MsgBox "Тучков"
'    End If
    If InStr(X.Cells(1, 3), X.Cells(1, 1)) = 0 Then
        Buddy = X.Cells(1, 1)
    Else
        Buddy = ""
    End If
    OwnerId = X.Cells(1, 4)
End Function
Function OppByPay(PayKod)
'
' получение имени Проекта по коду Платежа 1C из диста SF
'   17.2.2012
    On Error Resume Next
    OppByPay = ""
    OppByPay = WorksheetFunction.VLookup(PayKod, _
        Sheets(SF).Range("B2:K77777"), 10, False)
    On Error GoTo 0
End Function
Function OppNbyId(OppId) As Integer
'
' получение номера Проекта OppN по его Id в SF
'   21.3.2012
'   25.4.12 - Lookup case insensitive - заменен на цикл по SFopp
    
    Dim i
    OppNbyId = 0
    For i = 2 To EOL_SFopp
        If DB_SFDC.Sheets(SFopp).Cells(i, SFOPP_OPPID_COL) = OppId Then
            OppNbyId = DB_SFDC.Sheets(SFopp).Cells(i, SFOPP_OPPN_COL)
            Exit For
        End If
    Next i
End Function

Function OppNbyPay(PayKod) As Integer
'
' получение номера Проекта по коду Платежа 1C из диста SF
'   25.2.2012
    OppNbyPay = 0
    On Error Resume Next
    OppNbyPay = WorksheetFunction.VLookup(PayKod, _
        Sheets(SF).Range("B2:L77777"), 11, False)
    On Error GoTo 0
End Function
Function OppIdbyPay(PayKod)
'
' получение Id Проекта по коду Платежа 1C из диста SF
'   11.3.2012
    OppIdbyPay = ""
    On Error Resume Next
    OppIdbyPay = WorksheetFunction.VLookup(PayKod, _
        Sheets(SF).Range("B:S"), 18, False)
    On Error GoTo 0
End Function
Function OppIdbyContrId(ContrId)
'
' получение Id Проекта по коду Платежа 1C из диста SF
'   11.3.2012
    OppIdbyPay = ""
    On Error Resume Next
    OppIdbyPay = WorksheetFunction.VLookup(PayKod, _
        Sheets(SF).Range("B:S"), 18, False)
    On Error GoTo 0
End Function
Function OppOwner(PayKod)
'
' получение фамилии Продавца по коду Платежа из диста SF
'   15.2.2012
    On Error Resume Next
    OppOwner = ""
    OppOwner = WorksheetFunction.VLookup(PayKod, Sheets(SF).Range("B2:I77777"), 8, False)
    On Error GoTo 0
End Function
Function OppNameByN(OppN)
'
' получение названия Проекта из диста SFopp по его Id
'   23.3.2012
    Dim S As String
    S = OppN
    On Error Resume Next
    OppNameByN = ""
    OppNameByN = WorksheetFunction.VLookup(S, Sheets(SFopp).Range("B:C"), 2, False)
    On Error GoTo 0
End Function

Function ContrNbyPay(PayKod)
'
' получение номера Договора по коду Платежа 1C из диста SF
'   10.3.2012
    ContrNbyPay = ""
    On Error Resume Next
    ContrNbyPay = WorksheetFunction.VLookup(PayKod, _
        Sheets(SF).Range("B:Q"), 16, False)
    On Error GoTo 0
End Function
Function ContrN(ContrCod)
'
' получение номера Договора по коду Договора (/) из листа SFD
'   10.3.2012
    ContrN = ""
    On Error Resume Next
    ContrN = WorksheetFunction.VLookup(ContrCod, _
        Sheets(SFD).Range("B:C"), 2, False)
    On Error GoTo 0
End Function
Function ContractId(ContrCod) As String
'
' получение Id по имени контракта в форме <Осн.договор/Договор>
' если контракта нет в SF, возвращает ""
'   14.2.2012
'   28.2.2012 - изменение спецификации - используем <Осн.договор/Договор>

    ContractId = ""
    On Error Resume Next
    ContractId = WorksheetFunction.VLookup(ContrCod, _
        Sheets(SFD).Range("B:O"), 14, False)
    On Error GoTo 0
End Function
Function ContrCod(Contr, MainContr) As String
'
' Преобразование в форму <Осн.договор/Договор>
'   28.2.12
'  29.12.12 - Если в MainContr есть слово "Договор" - выбрасываем

    Dim C As String
    
    Contr = Trim(Contr)
    MainContr = Trim(MainContr)
    
    C = MainContr
    If C = "" Then
        C = Contr
    Else
        C = C & "/" & Contr
    End If
    C = Replace(C, "Договор", "")
    
    ContrCod = Trim(C)
End Function
Function ContrOppN(ContrId) As Integer
'
' проверка, связан ли Договор ContrId с Проектом. Если да - возвращает номер Проекта
'   28.2.12
'   23.3.12 - почему-то не работает Lookup - заменен на цикл по SFD
    
    Dim i
    ContrOppN = 0
    For i = 2 To EOL_SFD
        If Sheets(SFD).Cells(i, SFD_CONTRID_COL) = ContrId Then
            ContrOppN = Sheets(SFD).Cells(i, SFD_OPPN_COL)
            Exit For
        End If
    Next i
End Function
Function PayIdByK(PayK) As String
'
' - PayIdByK(PayK)  - получение Id SF по коду Платежа
'   13.5.2012

    PayIdByK = ""
    On Error Resume Next
    PayIdByK = WorksheetFunction.VLookup(PayK, _
        Sheets(SF).Range("B:R"), 17, False)
    On Error GoTo 0
End Function
Sub testPayInvByK()
    Call PayInvByK("542 от 16.09.09ЗАО ""ТЕКНИП РУС""218")
    Call PayInvByK("542 от 16.09.09ЗАО")                    ' не найдет!
    Call PayInvByK("549 от 31.01.2012ОАО ""КБСМ""63")
End Sub
Function PayInvByK(PayK) As Integer
'
' - PayInvByK(PayK)  - получение поля Номер Счета по коду Платежа
'   13.5.2012
'   24.5.12 - обрезаем Счет до вида Сч-12 от 01.02.11
    
    On Error Resume Next
    PayInvByK = InvoiceN(WorksheetFunction.VLookup(PayK, _
        Sheets(SF).Range("B:H"), 7, False))
    On Error GoTo 0
End Function
Function InvoiceN(Str) As Integer
'
' - InvoiceN(Str) возвращает число - номер Счета по текстовой строке Str
'   24.5.12

    Dim X As Integer        ' номер символа - пробела в Str
    Dim Y As Integer        ' позиция "Cч-" в Str
    Dim S As String         ' фрагмент Str с номером Счета
    
    InvoiceN = 0
    If Str = "" Then Exit Function
    X = WorksheetFunction.Search(" ", Str)
    Y = WorksheetFunction.Search("Сч-", Str)
    S = Mid(Str, Y + 3, X - Y - 3)
    
    If Not IsNumeric(S) Then Exit Function
    InvoiceN = S
End Function
Sub testInvoiceN()
    Call InvoiceN("""Сч-102 от 28.02.11 Валюта сч -рубль""")
End Sub
Function IsRightContrOppLink(OppId, ContrCode)
'
' возвращает True, если Проект OppId связан с Договором ContrCode
'   1.4.12

    Dim i
    
    IsRightContrOppLink = False
    If OppId = "" Or ContrCode = "" Then Exit Function
    
    With DB_SFDC.Sheets(SFD)
        For i = 2 To EOL_SFD
            If .Cells(i, SFD_OPPID_COL) = OppId Then
                If .Cells(i, SFD_COD_COL) = ContrCode Then
                    IsRightContrOppLink = True
                    Exit Function
                End If
            End If
        Next i
    End With
End Function
Function ContrOK(OppN, ContrK) As Boolean
'
' возвращает True если в SF Платеж действительно связан с Проектом и Договором как в 1С
'   8.3.12

    Dim DogOppN As Integer
    
    ContrOK = True
    If ContrK = "" Then Exit Function
    DogOppN = 0
    On Error Resume Next
    DogOppN = WorksheetFunction.VLookup(ContrK, _
        Sheets(SFD).Range("B:Q"), 16, False)
    On Error GoTo 0
    
    If DogOppN <> OppN Then ContrOK = False
End Function
