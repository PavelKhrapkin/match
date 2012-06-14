Attribute VB_Name = "Paint"
'-------------------------------------------------------------------------------
' Форматирование и раскрашивание листов Платежей и Договоров в MatchSF-1C.xlsm
'(*) PaymentPaint   - раскрашиваем и форматируем Лист Платежей из 1С
'(*) ContractPaint  - раскрашиваем и форматируем лист Договоров
' -  SFaccCol(SheetN, ResLines)  - раскраска колонки 1 по SFacc
'
' 27.5.2012
Option Explicit

Sub PaymentPaint()
Attribute PaymentPaint.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' Раскрашиваем Лист Платежей
'
' * кнопка <Часы> на отчете Платежей
'
'   20.4.2012
'   30.4.12 - SFaccColFill

    Const Reslines = 3  ' число строк в сводке под отчетом

    Lines = ModStart(1, "Раскрашиваем лист Платежей") - Reslines
     
'*** вначале всегда восстанавливаем обычный формат листа платежей
    SheetSort 1, 2          ' сортируем лист Платежей по порядку номеров в колонке В
    
    Worksheets(1).Columns("A:AC").Select
    Selection.EntireColumn.Hidden = False ' проявляем все невидимые (Hide) ячейки
    
    Range("J:Q,T:U,W:X").Select          ' делаем невидимыми ненужные колонки
    Selection.EntireColumn.Hidden = True 'валютных проводок, расходных кредитов,
                                         'отделов и фирм
' поправляем ширину колонок для удобства дальнейшей работы с SF
    Columns("G:G").ColumnWidth = 6.29   ' Дата прихода денег
    Columns("F:F").ColumnWidth = 11.57  ' Платежный документ
    Columns("H:H").ColumnWidth = 7.29   ' Счет
    Columns("Y:Y").ColumnWidth = 8.5    ' Договор
    Columns("Z:Z").ColumnWidth = 14     ' Основной Договор
    
' переводим Даты прихода денег из текстового формата
    Range("G2:G" & Lines).Select
    
    Selection.NumberFormat = "dd/mm/yy;@"
'    Selection.NumberFormat = "@"
    Columns("G:G").Select
    Selection.Replace What:=".", Replacement:=".", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

' выбираем в отчете отсутствующие в SF платежи и окрашиваем те, что надо заносить
    ActiveSheet.Range("A1:R" & Lines).AutoFilter Field:=4, Criteria1:="="
    
    With Range("R1:R" & Lines)
        .AutoFilter Field:=18, _
            Criteria1:=">30000", Operator:=xlAnd, Criteria2:="<=300000"
            With .Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent6
                .PatternTintAndShade = 0
                .pattern = xlSolid
            End With
        .Interior.TintAndShade = 0.599993896298105  ' от 30k до 300k- бледный
        .AutoFilter Field:=18
            
        .AutoFilter Field:=18, _
            Criteria1:=">300000", Operator:=xlAnd, Criteria2:="<=500000"
        .Interior.TintAndShade = 0.399975585192419  ' от 300k до 500k- бежевый
        .AutoFilter Field:=18

        .AutoFilter Field:=18, _
            Criteria1:=">500000", Operator:=xlAnd, Criteria2:="<1000000"
        .Interior.TintAndShade = -0.249977111117893 ' от 0,5 до 1 млн - оранжевым
        .AutoFilter Field:=18

        .AutoFilter Field:=18, Criteria1:=">=1000000"   ' платежи свяше 1 млн.р.
        .Interior.Color = 13209                         '     окрашиваем коричневым
    End With
    ActiveSheet.AutoFilterMode = False
    Rows("1:1").AutoFilter

    Pnt 4, 1, rgbLightGreen       ' Платежи занесенные в SF - зеленые
    Pnt 25, "<>", rgbLightBlue, 1 ' Договоры в колонке Y голубые
    Pnt 26, "<>", rgbLightBlue, 1 ' Осн.Договоры в колонке Z голубые
    Pnt 19, "=*Auto*", rgbADSK, 1 ' Товары Autodesk
    
    Call ADSK_P_Paint             ' раскраска Платежей с учетом SF_PA

    Call DateCol(1, 7)            ' преобразование колонки 7 в Дату
    
'    Rows("1:1").Interior.Color = RGB(255, 255, 255) 'первая строка без окраски
    
    Application.StatusBar = "Платежи " & Sheets(1).Name & " готовы к работе!"

End Sub
Sub ContractPaint()
Attribute ContractPaint.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' Раскрашиваем Лист Договоров
'   26.1.2012
'   1.3.12 -  добавлена колонка Поставщик
'  30.4.12 - SFaccColFill
'   1.5.12 - убрали ContrToDl и вообще занесение новых Договоров в SF

    Const Reslines = 7  ' количество строк в сводке внизу отчета
    
    Lines = ModStart(3, "Раскрашиваем лист Договоров")
    
'*** вначале всегда восстанавливаем обычный формат листа Договоров
    SheetSort 3, 4          ' сортируем лист Платежей по порядку номеров в колонке В
    
    ' все колонки листа Договоров - видимые
    ActiveSheet.UsedRange.EntireColumn.Hidden = False

    Rows("1:" & Lines).Font.Name = "Calibri"    ' весь отчет в формате Calibri 8
    Rows("1:" & Lines).Font.size = 8
    
    Range("T" & Lines - 5 & ":V" & Lines - 5).NumberFormat = "#,##0"
    
' поправляем ширину колонок для удобства дальнейшей работы с SF
    Columns("A:A").ColumnWidth = 1.86   ' SF Орг - Имя Организации в SF
    Columns("B:B").ColumnWidth = 7.86   ' Статус в SF
    Columns("C:C").ColumnWidth = 6.57   ' Тест Организации в SF <!> надо бы им заняться 2.1.12
    Columns("D:D").ColumnWidth = 6.57   ' Номер в SF
    Columns("E:E").ColumnWidth = 1.71   ' Получено - оплачено по Счету в 1С
    Columns("F:F").ColumnWidth = 1.57   ' Выставлены Счета
    Columns("G:G").ColumnWidth = 1.43   ' Отсканировано
    Columns("H:H").ColumnWidth = 0.75   ' <Основной>/<Договор>
    Columns("I:I").ColumnWidth = 0.75
    Columns("J:J").ColumnWidth = 8.45   ' Номер Договора в 1С
    Columns("K:K").ColumnWidth = 0.75
    Columns("L:L").ColumnWidth = 9      ' Дата подписания
    Columns("M:M").ColumnWidth = 7.57   ' наша фирма
    Columns("N:N").ColumnWidth = 10.29  ' Продавец
    Columns("O:O").ColumnWidth = 20.57  ' Заказчик в 1С
    Columns("P:P").ColumnWidth = 12.75  ' Поставщик
    Columns("Q:Q").ColumnWidth = 15.86  ' Тип договора
    Columns("R:R").ColumnWidth = 8.14   ' Основной Договор
    Columns("S:S").ColumnWidth = 5.57   ' Номер у заказчика
    Columns("T:T").ColumnWidth = 14     ' Сумма (план)
    Columns("U:U").ColumnWidth = 10     ' Сумма по счетам
    Columns("V:V").ColumnWidth = 10     ' Сумма по платежам
    Columns("W:W").ColumnWidth = 5.14   ' Вылюта
    Columns("X:X").ColumnWidth = 11.43  ' Счета и платежи
    Columns("Y:Y").ColumnWidth = 24.57  ' Примечание
    
    Lines = Lines - Reslines    ' В дальнейшем с сводками не работаем
    
    Pnt 2, "Закрыт", rgbLightGreen      ' Договоры Закрытые в SF - зеленые
    Pnt 2, "Открыт", rgbOrange          ' Открытые Договоры - оранжевые
    Pnt 2, "Черновик", rgbLightBlue     ' Черновики - голубые
    Pnt 2, "Не состоялся", Antique      ' Не состоялся - Antique
    Pnt 5, 1, LimeG, 1                  ' Оплаченные - темно зеленый
    Pnt 6, 1, rgbOlive, 1               ' Выставлен Счет - оливковый
    Pnt 7, 1, rgbViolet, 1              ' Выставлен Счет - оливковый
    
    ModEnd 3
End Sub
Sub TestAccFill()
    SFaccColFill PAY_SHEET  ' в колонке 1 если Организация есть в SF
    SFaccCol PAY_SHEET, PAY_RESLINES    ' раскрашиваем колонку A
End Sub
Sub SFaccColFill(SheetN)
'
' заполняем колонку 1 листа SheetN 1 если Организация в AccCol есть в SF
' это Case Sensitive для имен Организаций, в отличие от VLOOKUP
'   30.4.12
'  1.5.12 - проверка на пустое поле Клиент в Договоре
'  3.5.12 - использование Hash таблиц А.Пасс

    Dim L, i, j, AccCol As Integer
    Dim Acc, MSG As String
    Dim accInSF As String
    Dim t0, t1
    
    Static accHTable(0 To 5000) As String
    Static accVTable(0 To 5000) As String
    
    hashInit accHTable, accVTable
    
    t0 = Timer
    
    MSG = "SFaccColFill -- заполнение колонки 1 для листа " & SheetN
    
    ModStart SheetN, MSG, True
    
    Select Case SheetN
        Case PAY_SHEET:
            L = EOL_PaySheet
            AccCol = PAYACC_COL
        Case DOG_SHEET:
            L = EOL_DogSheet
            AccCol = DOG1CACC_COL
        Case Acc1C:
            L = EOL_Acc1C
            AccCol = A1C_NAME_COL
        Case Else
            MsgBox "неправильный лист в SFaccColFill", , "FATAL ERROR"
            Stop
    End Select
    
 ' -- проход по листу Организаций в SF SFacc для каждой строки SheetN
    With Sheets(SFacc)
        For i = 2 To L
            Progress i / L
 
            Acc = Replace(Compressor(Sheets(SheetN).Cells(i, AccCol)), vbCrLf, "")
            If Acc = "" Then
                accInSF = ""
            Else
                accInSF = hashGet(accHTable, accVTable, Acc)
            End If
            If accInSF = "$" Then
                Sheets(SheetN).Cells(i, 1) = ""
                If Acc <> "" Then
                    For j = 2 To EOL_SFacc
                        If Acc = Replace(Sheets(SFacc).Cells(j, SFACC_ACC1C_COL), "" & vbCrLf, "") Then
                            Sheets(SheetN).Cells(i, 1) = "1"
                            Exit For
                        End If
                    Next j
                End If
                hashSet accHTable, accVTable, Acc, Sheets(SheetN).Cells(i, 1)
            Else
                Sheets(SheetN).Cells(i, 1) = accInSF
            End If
        Next i
    End With

    t1 = Timer - t0
'    MsgBox "Time =" & t1

    ModEnd SheetN
End Sub
Sub SFaccCol(SheetN, Optional Reslines As Integer = 0)
'
' Раскрашиваем колонку А:А в листе SheetsN по SFacc - справочнику Организаций
' Если данная Организация есть в SF (значение в колонке =1)она желтая, иначе красная
'   12.1.2012
'   25.1.2012 - исправлена ошибка с Lines; теперь используется Private L
'   29.2.2012 - bug fix, EOL вместо AutoFulterReset
'   13.5.2012 - сброс прежней раскраски листа

    Dim L
    
    L = EOL(SheetN) - Reslines
    
    Range("B2:AZ" & L).Interior.Color = rgbWhite    ' сбрасываем старую раскраску листа
    Range("A2:A" & L).Interior.Color = rgbYellow
    Range("A1:A" & L).AutoFilter Field:=1, Criteria1:="="
    Range("A2:A" & L).Interior.Color = rgbRed  ' неизвестные SF Организации в колонке 1 красные
    ActiveSheet.UsedRange.AutoFilter Field:=1
    
    Rows("1:1").RowHeight = 45          ' первый ряд тройной высоты
    Rows("2:" & L).RowHeight = 15       ' высота строк до конца = 15
    
End Sub
Sub ADSK_P_Paint()
'
' - ADSK_P_Paint()  - раскраска Платежей ADSK, с учетом наличия связи SF_PA
'   18.5.12

    Dim i As Long
    Dim j As Long
    
    With Sheets(PAY_SHEET)
        For i = 2 To EOL_PaySheet
            If .Cells(i, PAYINSF_COL) = 1 Then
                If GoodType(.Cells(i, PAYGOOD_COL)) = WE_GOODS_ADSK Then
                    If IsP_AbyN(i - 1) Then
                        .Cells(i, PAYGOOD_COL).Interior.Color = rgbPink
                    End If
                End If
            End If
        Next i
    End With
    
End Sub

Function IsP_AbyN(Nstr) As Boolean
'
' - IsP_AbyN(Nstr)  - возвращает TRUE, если строка Nstr Платежа связзана с ADSK
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
