Attribute VB_Name = "Paint"
'---------------------------------------------------------------------------
' Загрузка и раскраска отчетов из 1С и SalesForce.com
'
' S PaymentPaint()  - Раскрашиваем Лист Платежей 1C и восстанавливаем пятку
' S ContractPaint() - Раскрашиваем Лист Договоров
' - Paint(iStr,Col,Criteria,Color,[Mode]) - раскраска ячейки (iStr,Col) в цвет Color
' - Acc1C_Bottom()  - перенос первыx трех строк Acc1С в пятку
' S AccPaint()      - окраска колонки А - Организация есть в SF
' S SF_Paint()      - окраска колонки А отчета SF по Платежам
' ? IsP_AbyN(Nstr)  - возвращает TRUE, если строка Nstr Платежа связана с ADSK
'
' 8.11.2012 П.Л.Храпкин match 2.0
' 26.8.13 - ревизия для match2.1

Option Explicit
Sub PaymentPaint()
'
' S PaymentPaint()  - Раскрашиваем Лист Платежей 1C и восстанавливаем пятку
'
' 24.6.12 переписано для match 2.0
'  7.8.12 оформлено как Шаг
' 31.8.12 - внедрение StepIn
'  7.2.13 - параметр BottomHDR; окраска всей строки, занесенной в SF
' 18.8.13 - стираем строки нал
' 24.8.13 - восстанавливаем пятку, если она есть

    StepIn

    Dim i As Integer
    Dim Rub, Doc            'поля "Итого руб" и "Плат.док"
    Dim Fsummary As String  'имя пятки Платежей в TOC
    
    RepTOC.EOL = EOL(RepTOC.Name)
    Range("A1:AC" & RepTOC.EOL).Interior.Color = rgbWhite   ' сбрасываем окраску
    Rows("2:" & RepTOC.EOL).RowHeight = 15    ' высота строк до конца = 15
    
    With DB_1C.Sheets(PAY_SHEET)
        i = 2
        Do While i <= RepTOC.EOL
            Progress i / RepTOC.EOL
            If .Cells(i, PAYINSF_COL) = 1 Then          ' зеленые Платежи в SF
                Range(Cells(i, 2), Cells(i, .Columns.Count)).Interior.Color = rgbLightGreen
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
            
'-- стираем нал
            Doc = Trim(.Cells(i, PAYDOC_COL))
            If Doc = "" Or InStr(Doc, "авт нал") <> 0 Then
                .Rows(i).Delete
                i = i - 1
                RepTOC.EOL = RepTOC.EOL - 1
            End If
            i = i + 1
        Loop
        
'-- переписываем пятку, если она есть - вслед за EOL должна быть пустая строка
        RepTOC.EOL = EOL(RepTOC.Name, Workbooks(RepTOC.RepFile))
        With Workbooks(RepTOC.RepFile).Sheets(RepTOC.SheetN)
            For i = 1 To PAYGOODTYPE_COL
                If Trim(.Cells(RepTOC.EOL - 1, i)) <> "" Then GoTo AddSummary
            Next i
            .Rows(RepTOC.EOL + 1 & ":" & RepTOC.EOL + 2).Delete 'пятка есть- стереть
            RepTOC.EOL = RepTOC.EOL - 2
            
AddSummary: Fsummary = DB_MATCH.Sheets(TOC).Cells(RepTOC.iTOC, TOC_FORMSUMMARY)
            DB_MATCH.Sheets(Header).Range(Fsummary).Copy _
                Destination:=.Cells(RepTOC.EOL + 2, 1)
        End With
    End With
End Sub
Sub ContractPaint()
'
' S ContractPaint() - Раскрашиваем Лист Договоров
' 10.8.12
'  1.9.12 - StepIn
' 14.9.12 - раскраска с Paint
'  4.11.12 - флаг "Не отсканировано" - красный

    StepIn
    
    Dim D As TOCmatch
    D = GetRep(DOG_SHEET)
    Dim i As Long
    
    ScreenUpdate False
    
    For i = 2 To D.EOL
        ActiveSheet.Rows(i).Interior.Color = rgbWhite
        Paint i, DOGSFSTAT_COL, "Закрыт", rgbLightGreen ' Договоры Закрытые в SF- зеленые
        Paint i, DOGSFSTAT_COL, "Открыт", rgbOrange     ' Открытые Договоры - оранжевые
        Paint i, DOGSFSTAT_COL, "Черновик", rgbLightBlue ' Черновики - голубые
        Paint i, DOGSFSTAT_COL, "Не состоялся", Antique ' Не состоялся - Antique
        Paint i, DOGSFSTAT_COL, "Нет в SF", rgbWhite    ' Нет в SF - не окрашиваем
        Paint i, DOGPAID1C_COL, "1", LimeG, 1           ' Оплаченные - темно зеленый
        Paint i, DOGISINV1C_COL, "1", rgbOlive, 1       ' Выставлен Счет - оливковый
        Paint i, DOG1CSCAN_COL, "1", rgbViolet, 1       ' Отсканировано - фиолетовый
        Paint i, DOG1CSCAN_COL, "0", rgbRed, 1          ' НЕ Отсканировано - красный
    Next i
    
    ScreenUpdate True
    
''''-- копируем пятку в Договоры
'''    DB_MATCH.Sheets(Header).Range("HDR_1C_Contract_Summary").Copy _
'''            Destination:=ActiveSheet.Cells(D.EOL + 1, 1)
            
End Sub
Sub Paint(iStr As Long, Col As Long, Criteria As String, Color, Optional Mode As Integer = 0)
'
' - Paint(iStr,Col,Criteria,Color,[Mode]) - раскраска ячеки (iStr,Col) в цвет Color
'                            при значении Criteria, или вся строка если указано Mode=1
' 14.9.12

    Const DOG_COLS = 26     ' число колонок в таблице Договоров
    
    With ActiveSheet
        If .Cells(iStr, Col) = Criteria Then
            If Mode = 1 Then
                .Cells(iStr, Col).Interior.Color = Color
            Else
                Range(Cells(iStr, 2), Cells(iStr, DOG_COLS)).Interior.Color = Color
            End If
        End If
    End With
End Sub
Sub Acc1C_Bottom()
'
' - Acc1C_Bottom() - перенос первыx трех строк Acc1С в пятку и в Шаблон пятки
'   14.8.12
'   26.8.13 - добавлен AutoFilterReset и обновляю Шаблон пятки из отчета Acc1C

    Dim R As TOCmatch
    Dim b As Range, FF As Range
    Dim Fsummary As String
    
    R = GetRep(Acc1C)
    DB_1C.Sheets(Acc1C).Activate
    With DB_MATCH
        Fsummary = .Sheets(TOC).Cells(R.iTOC, TOC_FORMSUMMARY)
        Set FF = .Sheets(Header).Range(Fsummary)
    End With
    With ActiveSheet
        Set b = .Range(.Cells(1, 1), .Cells(2, 1))
        b.Copy Destination:=FF.Cells(1, 3)
        Set b = .Rows("1:3")
        b.Copy Destination:=.Cells(R.EOL + 2, 1)
        b.Delete
    End With
    AutoFilterReset R.Name
End Sub
Sub AccPaint()
'
' S AccPaint() - окраска колонки А - Организация есть в SF
'   14.8.12
'   31.8.12 - внедрение StepIn

    StepIn

    Dim i As Long
    Dim RepTo As TOCmatch
    Dim R As Range
    
    ScreenUpdate False
    
    RepTo = GetRep(ActiveSheet.Name)
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
    
    ScreenUpdate True
End Sub
Sub SF_Paint()
'
' S SF_Paint() - окраска колонки А отчета SF по Платежам
'
'   25.8.13

    StepIn
    
    Dim i As Long, R As TOCmatch
    R = GetRep(ActiveSheet.Name)
    R.EOL = EOL(R.SheetN) - R.ResLines
    
    With Workbooks(R.RepFile).Sheets(R.SheetN)
        For i = 2 To R.EOL
            Progress i / R.EOL
            If .Cells(i, 1) = "" Then
                If InStr(.Cells(i, SF_SAIL_COL), "Лидер") Then
                    .Cells(i, 1).Interior.Color = vbBlue
                Else
                    .Cells(i, 1).Interior.Color = vbRed
                End If
            End If
        Next i
    End With
End Sub

Function IsP_AbyN(Nstr) As Boolean
'
' - IsP_AbyN(Nstr)  - возвращает TRUE, если строка Nstr Платежа связана с ADSK
'   18.5.2012

    Dim SFid As String
    Dim ADSKid As String
    Dim i As Long

    IsP_AbyN = False
    SFid = "": ADSKid = ""
    On Error Resume Next
    SFid = WorksheetFunction.VLookup(Nstr, _
        Sheets(SF).Range("A:R"), 18, False)
    On Error GoTo 0
    If SFid = "" Then Exit Function
    For i = 2 To EOL_SFlnkADSK
        If Sheets(SF_PA).Cells(i, 2) = SFid Then
            IsP_AbyN = True
            Exit Function
        End If
    Next i
End Function
