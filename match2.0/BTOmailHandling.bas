Attribute VB_Name = "BTOmailHandling"
'------------------------------------------------------------------------------------
' BTOhandling - обработка e-mail'ов от CSD по отгрузке Обновлений SN на Склад
'
'   6.4.13
'
'[*] BTO_Mail_track()   - чтение и обработка файла BTOmails
' -  BTOmailHandle (SN, BTOmsg, BTOmsgLines) - обработка письма БТО
' -  IsSNonStock(SN)    - return TRUE if SN is registered on Stock

Option Explicit

Sub BTO_Mail_track()
'
'[*] BTO_Mail_track() - чтение и обработка файла BTOmails
'
' When string in file contains BTOstamp, read mail - seek SN on Stock
'   12.6.12
'   15.6.12 - иногда строка Autodesk переносится. Просматриваем две,
'             чтобы не потерять SN
'   18.6.12 - проверка Счета CSD по листу Заказов
'   14.11.12 - перенос в match 2.0
'    6.04.13 - открываем через TOCmatch и размещаем в Stock.xlsx, переписан код
'
'------ INITIALIZATION AND LOCAL DECLARATION SECTION ---------------------
'''''    Const BTOfileName = "BTOmails.txt"  ' входной файл - письма из Outlook
    Dim R As TOCmatch       'входной файл BTOmails для обработки
    Dim iMail As Integer    '= число обработаных мейлов
'''    Dim iSN As Integer      '= число SN, не проведенных по Складу
'''    Dim iADSK As Integer    '= номер строки ADSK из файла BTOmails.txt
    Dim MailDate As Date    '= дата и время письма БТО в mail Subject
    Dim CSD_Inv As String   '= Заказ/Счет CSD
    Dim Descr As String     'Описание продукта Autodesk
    Dim SN As String        '= текущий SN
    Dim Seats As Long       'Количество мест
    Dim i As Long, S As String
    
    StepIn
    R = GetRep(BTOmails)
    
'---------------------- CODE SECTION -----------------------------------
    With Workbooks(R.RepFile).Sheets(R.SheetN)
        For i = 1 To R.EOL
            Progress i / R.EOL
            S = .Cells(i, 1)
            If InStr(S, "БТО: Обновление по подписке") <> 0 Then
                iMail = iMail + 1
                MailDate = Mid(S, 2, WorksheetFunction.FindB("]", S) - 2)
            ElseIf InStr(S, "Счет:#") <> 0 Then
                CSD_Inv = Mid(S, 7)
            ElseIf InStr(S, "Auto") Then
'''                SN = RemIgnoredSN()
                BTOmailHandle MailDate, CSD_Inv, Descr, SN, Seats
            ElseIf InStr(S, "------") Then
                CSD_Inv = "": MailDate = 0
            End If
        Next i
    End With
'----------------------- SUMMARY SECTION -------------------------------
'''    Columns("A:J").Select           ' текст без WrapText
'''    Selection.WrapText = False
'''
'''    MS "В файле " & BTOmails & " просмотрено " _
'''        & Str$(iMail) & " писем БТО, для " & Str$(iSN) _
'''        & " из них проводок по Складу не найдено"
''''''    ModEnd STOCK_SHEET
End Sub
Function BTOmailHandle(MailDate As Date, CSD_Inv As String, Descr As String, SN As String, Seats As Long) As Boolean
'''''Function BTOmailHandle(SN, BTOmsg, BTOmsgLines) As Boolean
'
' - BTOmailHandle (SN, BTOmsg, BTOmsgLines) - обработка письма БТО
'       возвращает FALSE, если письмо обработать не удалось или
'       указанный в нем SN был проведен по Складу и обрабатывать не надо
'   12.6.12
'   15.6.12 - добавлены колонки "Доставка со Склада", "Дата оплаты" и "Счет 1С"

    Dim Sale As String      'поле BTO "Продавец"
    Dim Client As String    'поле BTO "Заказчик"
    Dim Delivery As String  'поле "Доставка со склада"
    Dim Paid As String      'поле "Дата оплаты Счета в 1С"
    Dim Inv1C As String     'поле "Счет 1С"
    Dim iStock As Integer   '= номер строки по Складской книге
    Dim iCSD As Integer     '= по листу Заказов
    Dim SN_SF As SNatr      '= структура SN в SF
    Dim iSF As Integer      '= номер строки в отчете ADSKfrSF по SN
    Dim GoodADSK As String  'поле ВТО "Товар ADSK" - строка из письма
    
    Dim Msg As String
    Const MaxStrLen = 100
    Dim i As Integer
    
    Const InvCSD = 37   'позиция счета CSD в строке Subject после даты
'    Dim iCSD As Integer '= номер символа - счета CSD в Subject
    
    EOL_BTO = EOL_BTO + 1
    
'---- выделение даты и времени письма, атрибутов CSD из Subject
    With Sheets(BTO_SHEET)
        For i = 2 To MaxStrLen
            If Mid(BTOmsg(1), i, 1) = "]" Then
                .Cells(EOL_BTO, BTO_DATE_COL) = Mid(BTOmsg(1), 2, i - 2)
                iCSD = i + InvCSD
                .Cells(EOL_BTO, BTO_CSDATR_COL) = _
                    Trim(Mid(BTOmsg(1), iCSD, Len(BTOmsg(1)) - iCSD - 7))
                Exit For
            End If
        Next i
        
'---- работа с полным текстом письма БТО
        Msg = ""
        For i = 1 To BTOmsgLines
            If Not InStr(BTOmsg(i), "-------") <> 0 Then
                Msg = Msg & BTOmsg(i) & vbCrLf
            End If
        Next i
        .Cells(EOL_BTO, BTO_MAIL_COL) = Msg

'---- работа со строкой "Товар ADSK"
        For i = 2 To BTOmsgLines
            If InStr(BTOmsg(i), "Auto") <> 0 Then
                .Cells(EOL_BTO, BTO_GOOD_COL) = _
                    Mid(BTOmsg(i), 3, Len(BTOmsg(i)) - 3)
                Exit For
            End If
        Next i
'---- работа с Заказом через CSD
        If IsCSDinv(.Cells(EOL_BTO, BTO_CSDATR_COL), iCSD) Then
            With Sheets(OrderList)
                Paid = .Cells(iCSD, OL_PAIDDAT_COL)
                Inv1C = .Cells(iCSD, OL_INV1C_COL)
            End With
        Else
            Paid = "": Inv1C = ""
        End If
    
'---- работа с SN
        .Cells(EOL_BTO, BTO_SN_COL) = SN
               
        If Len(SN) <> 12 Then
            Sale = "<-- Нет SN в письме БТО -->"
            Client = "": Delivery = "": Sale = "": Inv1C = "": Paid = ""
        Else
            If IsSNonStock(SN, iStock) Then
               Delivery = Sheets(STOCK_SHEET).Cells(iStock, STOCK_DELIVERY_COL)
            End If
'---- SN из SF
            SN_SF = SNinSFatr(SN, iSF)  '<<< находим SN в SF >>>
            If SN_SF.ErrFlag Then
                Sale = "<-!- отсутствует в SF -!->"
                Client = ""
                ErrMsg TYPE_ERR, "В SF нет SN=" & SN
            Else
                
'---- запись в BTOlog
                With Sheets(ADSKfrSF)
                    Sale = .Cells(iSF, SFADSK_SALE_COL)
                    Client = .Cells(iSF, SFADSK_ACC1C_COL)
                End With
            End If
        End If
        .Cells(EOL_BTO, BTO_DELIVERY_COL) = Delivery
        .Cells(EOL_BTO, BTO_PAID_DATE_COL) = Paid
        .Cells(EOL_BTO, BTO_INV_1C_COL) = Inv1C
        .Cells(EOL_BTO, BTO_SALE_COL) = Sale
        .Cells(EOL_BTO, BTO_CLIENT_COL) = Client
    End With
End Function
Function IsSNonStock(SN, iStock) As Boolean
'
' - IsSNonStock(SN, iStock)    - return TRUE if SN is registered on Stock
'   11.6.12
'   15.6.12 возвращает номер строки по Складу
    
    Dim i As Integer
    
    IsSNonStock = False
    If SN = "" Then Exit Function
    With Sheets(STOCK_SHEET)
        For i = 2 To EOL_Stock
            If InStr(.Cells(i, STOCK_SN_COL), SN) <> 0 Then
'                Client = .Cells(i, STOCK_CLIENT_COL)
'                Dat = .Cells(i, STOCK_DATE_COL)
                IsSNonStock = True
                iStock = i
                Exit Function
            End If
        Next i
    End With
End Function
Function IsCSDinv(Str, iCSD) As Boolean
'
' - IsCSDinv(Str, iCSD) - возвращает TRUE и номер строки,
'              если номер заказа найден среди Заказов CSD
'   17.6.12

    Dim Inv As String   'поле "№ счета CSD" Заказов
    Dim MsgInv As String '= Номер счета СЫВ в письме БТО
    Dim S() As String   '= массив для разбора Str
    Dim Dat As Date     'поле "Дата счета CSD" Заказов
    Dim MsgDat As Date  '= "Дата счета CSD" по письму БТО
    
    IsCSDinv = False
    If Str = "" Then Exit Function

    S = Split(Str, " ")
    If S(1) <> "от" Then Exit Function
    MsgInv = S(0)
    MsgDat = S(2)
    
    With Sheets(OrderList)
        For iCSD = 2 To EOL_OrderList
If iCSD >= 437 Then
S = S
End If
            Inv = .Cells(iCSD, OL_CSDINVN_COL)
            Dat = .Cells(iCSD, OL_CSDINVDAT_COL)
            If Inv = MsgInv And Abs(Dat - MsgDat) < 5 Then
                IsCSDinv = True
                Exit Function
            End If
        Next iCSD
    End With
End Function
