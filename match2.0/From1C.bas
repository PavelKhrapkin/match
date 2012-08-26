Attribute VB_Name = "From1C"
'---------------------------------------------------------------------------
' Макросы для загрузки отчетов из 1С
'
'<*> From1Cpayment  - заменяет лист отчета из 1С "Приход денег на счета"
' -  SFmatchFill(SheetN)  - заполнение связей листа SheetN по SFDC
' -  CSmatch(Val,Col,[SheetN],[DB]) - Case Sensitive match возвращает номер строки
'           с Val в колонке Col листа SheetN в DB. Если Val не найден- возвращает 0.
'<*> From1Cdogovor  - заменяет лист отчета 1С "Договоры" новым из 1С
'<*> From1Caccount  - заменяет лист отчета 1С "Клиенты .." новым из 1С
'(*) FromStock      - замена Складской Книги в листе Stock
'
' 10.8.2012 П.Л.Храпкин match 2.0

Option Explicit
Sub SFlnkFill(DocFr, ColFr, ColFrId, ColVal, ColTo)
'
' - SFlnkFill(RepFr,ColFr,ColFrId, ColVal, ColTo) - "сшивает" отчет в DocFr с DocTo,
'       записывая в колонку ColTo Id рекорда найденного по значению в ColFr
' 8.8.12

    Dim DocTo As String ' имя входного Документа - отчета
    Dim RepFr As TOCmatch, RepTo As TOCmatch
    Dim Val
    Dim i As Integer, N As Integer
    
    DocTo = ActiveSheet.Name
    RepTo = GetRep(DocTo)
    Workbooks(RepTo.RepFile).Sheets(RepTo.SheetN).Activate
    RepFr = GetRep(DocFr)
    With Workbooks(RepTo.RepFile).Sheets(RepTo.SheetN)
        Workbooks(RepFr.RepFile).Sheets(RepFr.SheetN).Select
        For i = 2 To RepTo.EOL
            Progress i / RepTo.EOL
            Val = .Cells(i, ColVal)
            N = CSmatch(Val, ColFr)
            If N > 0 Then
                .Cells(i, ColTo) = Workbooks(RepFr.RepFile).Sheets(RepFr.SheetN).Cells(N, ColFrId)
            Else
                .Cells(i, ColTo) = ""
            End If
        Next i
    End With
End Sub
Sub DogStatus()
'
' - DogStatus()     - вставляет в лист Договоров статус из SND
'   11.8.12

    Dim i As Long, N As Long
    
    PublicProcessName = ""
    GetRep SFD
    GetRep DOG_SHEET
    DB_SFDC.Sheets(SFD).Activate
    With DB_1C.Sheets(DOG_SHEET)
        For i = 2 To RepTOC.EOL
            Progress i / RepTOC.EOL
            N = CSmatch(.Cells(i, DOGCOD_COL), SFD_COD_COL)
            If N > 0 Then
                .Cells(i, DOGSFSTAT_COL) = _
                    DB_SFDC.Sheets(SFD).Cells(N, SFD_STATUS_COL)
            End If
        Next i
    End With
End Sub
Sub ContractPaint()
'
' - ContractPaint() - Раскрашиваем Лист Договоров
' 10.8.12

    GetRep DOG_SHEET
    Call AutoFilterReset(DOG_SHEET)
    Rows(1).RowHeight = 50
    
    Pnt DOGSFSTAT_COL, "Закрыт", rgbLightGreen      ' Договоры Закрытые в SF- зеленые
    Pnt DOGSFSTAT_COL, "Открыт", rgbOrange          ' Открытые Договоры - оранжевые
    Pnt DOGSFSTAT_COL, "Черновик", rgbLightBlue     ' Черновики - голубые
    Pnt DOGSFSTAT_COL, "Не состоялся", Antique      ' Не состоялся - Antique
    Pnt DOGSFSTAT_COL, "Нет в SF", rgbWhite         ' Нет в SF - не окрашиваем
    Pnt DOGPAID1C_COL, 1, LimeG, 1                  ' Оплаченные - темно зеленый
    Pnt DOGISINV1C_COL, 1, rgbOlive, 1              ' Выставлен Счет - оливковый
    Pnt DOG1CSCAN_COL, 1, rgbViolet, 1              ' Отсканировано - фиолетовый
'-- копируем пятку в Платежи1С
    Range("Contract_Summary").Copy Destination:=ActiveSheet.Cells(RepTOC.EOL + 1, 1)
End Sub
Sub Acc1C_Bottom()
'
' - Acc1C_Bottom() - перенос первыx трех строк Acc1С в пятку
'   14.8.12

    Dim R As TOCmatch
    Dim b As Range
    
    R = GetRep(Acc1C)
    DB_1C.Sheets(Acc1C).Activate
    Set b = ActiveSheet.Rows("1:3")
    b.Copy Destination:=Cells(R.EOL + 2, 1)
    b.Delete
End Sub
Sub AccPaint()
'
' - AccPaid() - окраска колонки А - Организация есть в SF
'   14.8.12

    Dim i As Long
    Dim RepTo As TOCmatch
    Dim R As Range
    
    RepTo = GetRep(ActiveSheet.Name)
    DB_1C.Sheets(RepTo.SheetN).Activate
    With Workbooks(RepTo.RepFile).Sheets(RepTo.SheetN)
        For i = 2 To RepTo.EOL
            Progress i / RepTo.EOL
            Set R = .Cells(i, PAYISACC_COL)
            If R <> "" Then
                R.Interior.Color = rgbYellow
            Else
                R.Interior.Color = rgbRed
            End If
        Next i
    End With
End Sub

Sub testCSmatch()
    If "G" = "g" Then Stop
    Dim A
    ThisWorkbook.Sheets("Sheet1").Select
    A = CSmatch("g12", 1)
    A = CSmatch("g121", 1)
    
    ModStart REP_1C_P_PAINT
    Set DB_SFDC = Workbooks.Open(F_SFDC, UpdateLinks:=False, ReadOnly:=True)
    DB_SFDC.Sheets(SFacc).Select
    A = CSmatch("ОАО ""ЭХО""", 2)
    ModEnd
End Sub
Sub PaymentPaint()
'
' - PaymentPaint() - Раскрашиваем Лист Платежей 1C
' 24.6.12 переписано для match 2.0
'  7.8.12 оформлено как Шаг

    Dim i As Integer
    Dim Rub, Doc    'поля "Итого руб" и "Плат.док"
    
    PublicStepName = ""
    GetRep PAY_SHEET
    DB_1C.Sheets(PAY_SHEET).Select
    
    Range("A1:AC" & RepTOC.EOL).Interior.Color = rgbWhite   ' сбрасываем окраску
    Rows("2:" & RepTOC.EOL).RowHeight = 15    ' высота строк до конца = 15
    
    With Sheets(PAY_SHEET)
        For i = 2 To RepTOC.EOL
            Progress i / RepTOC.EOL
            If .Cells(i, PAYINSF_COL) = 1 Then          ' зеленые Платежи в SF
                Range(Cells(i, 2), Cells(i, AllCol)).Interior.Color = rgbLightGreen
            ElseIf Trim(.Cells(i, PAYDOC_COL)) = "" Or Trim(.Cells(i, PAYSALE_COL)) = "" Then
                .Cells(i, 1).EntireRow.Hidden = True    ' нал убираем
            Else
'-- окраска еще не занесенных Платежей в зависимости от суммы
                Rub = .Cells(i, PAYRUB_COL)
                If Rub >= 1000000 Then
                    .Cells(i, PAYRUB_COL).Interior.Color = rgbBrown
                ElseIf Rub > 500000 Then
                    .Cells(i, PAYRUB_COL).Interior.Color = rgbOrange
                ElseIf Rub > 300000 Then
                    .Cells(i, PAYRUB_COL).Interior.Color = rgbBisque
                ElseIf Rub > 30000 Then
                    .Cells(i, PAYRUB_COL).Interior.Color = rgbBeige
                End If
            End If
            
'-- окраска Договоров
            If .Cells(i, PAYDOGOVOR_COL) <> "" Then     ' голубые Договоры
                .Cells(i, PAYDOGOVOR_COL).Interior.Color = rgbLightBlue
            End If
            If .Cells(i, PAYOSNDOGOVOR_COL) <> "" Then  ' голубые Осн.Договоры
                .Cells(i, PAYOSNDOGOVOR_COL).Interior.Color = rgbLightBlue
            End If
'-- окраска товаров ADSK в соответствии с SF_PA
            If InStr(.Cells(i, PAYGOOD_COL), "Auto") <> 0 Then  ' Autodesk
                If .Cells(i, PAYADSK_COL) = "" Then
                    .Cells(i, PAYGOOD_COL).Interior.Color = rgbADSK
                Else
                    .Cells(i, PAYGOOD_COL).Interior.Color = rgbPink
                End If
            End If
'-- окраска колонки А - Организация есть в SF
''            If .Cells(i, PAYISACC_COL) = "1" Then
''                .Cells(i, PAYISACC_COL).Interior.Color = rgbYellow
''            Else
''                .Cells(i, PAYISACC_COL).Interior.Color = rgbRed
''            End If
            
'-- скрываем нал
            Doc = Trim(.Cells(i, PAYDOC_COL))
            If Doc = "" Or InStr(Doc, "авт нал") <> 0 Then .Rows(i).Hidden = True
            
        Next i
'-- копируем пятку в Платежи1С
        Range("Payment_Summary").Copy Destination:=.Cells(RepTOC.EOL + 1, 1)
    End With
'    ModEnd REP_1C_P_PAINT
End Sub
