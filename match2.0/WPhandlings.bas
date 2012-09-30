Attribute VB_Name = "WPhandlings"
'====================================================================
' Модуль для работы с листом рабочим листом WP
'   1.10.2012
'
' после выбора данных в Match они обрабатываются на рабочем листе WP

    Option Explicit    ' Force explicit variable declaration
Sub WP_PaidOpp(Form As String)
'
'
'

End Sub




''    Const WP = "WP"     ' имя рабочего листа
''    Const We = "We"     ' неизменные сведения о нашей организации и Продавцах
''    Const WPheader = "WPheader" ' лист - шаблон для WP
''
''    Dim PaidMod         ' режим вывода Платежей
''    Dim OppMod          ' режим вывода Проектов
''    Dim ContrMod        ' режим вывода Договоров
''    Dim ADSKmod         ' режим вывода объектов Autodesk
''
''    Dim Prow            ' позиция выбранная на листе Платежей
''    Dim Pcol            ' номер колонки на листе Платежей
''    Dim Wppay0, WPpayN, Wpopp0, WPoppN, Wpdod0, WPdonN
''    Dim EOL_WP
    
Sub AccPaid_Click()
'
' По кнопке <WP> на экране Платежей работа по выбранной трансакции
' на листе WP, в частности:
'   - если Организации нет в SF - распознавание и загрузка по 1С
'   - обзор Платежей Организации (занесенных и не занесенных в SF)
'   - обзор Проектов Организации (завершенных и незавершенных)
'   - обзор Договоров (только Черновиков и Открытых и всех - в SF и 1С)
'   - обзор и ревизия продуктов ADSK (с действующей подпиской и всех)
'
'   17/1/12
'   10/2/12 - использование WPform

    Dim Client1C As String
    
    Prow = ActiveCell.Row           ' Номер выбранной строки
    Pcol = ActiveCell.Column        ' Выбранная колонка
    Client1C = Range("I" & Prow)    ' выбранная Организация
    
    ModStart WP, "Работа с WP -- обдумывание и обработка Платежей из 1С"
    CheckSheet 1, 1, 6, "Плат. док."
    CheckSheet 1, 1, 7, "Дата прих. денег"
    
    If Prow > Sheets(1).UsedRange.Rows.Count - 3 Or _
       Pcol > Sheets(1).UsedRange.Columns.Count Then
        MsgBox "Выбор за пределами Платежей!", , "ERROR!"
        Exit Sub
    End If

    Call WPaid(-2)           ' загрузка Платежей по Организации на WP
    
    If Sheets(1).Cells(Prow, 1) = "" Then
        AccAddFrom1C Client1C   ' Выбранной Организации нет в SF - добавляем
        Exit Sub
    End If
    
    Wopp Client1C   ' вывод Проектов в режиме OppFlag
    Wcontr Client1C ' вывод Договоров в режиме ContrFlag
    WADSK Client1C  ' вывод имеющихся в Организации продуктов Autodesk в режиме ADSKFlag
    
    WPform.Show
    ModEnd WP
End Sub
Sub WPaid(Optional PaidMod)
'
' Вывод Платежей по выбранной Организации на листе WP
'   PaidMod<0   выводить с очисткой листа WP в режиме -PaidMod
'   PaidMod=1   выводить только выбранный платеж
'   PaidMod=2   выводить только платежи, которые есть в SF (зеленые)
'   PaidMod=3   выводить все Платежи по выбранной Организации
' на листе WP есть все Платежи (2), но ненужные невидимы
'
'   19.1.2012
'    4.2.2012 - Optional Public PaidMod. Если параметр указан -
'               переопределяет Public PaidMod
    
    Static FrL, ToL As Integer  ' диапазон номеров строк с Платежами на листе WP
    Dim P_Lines As Integer      ' число строк- Платежей листа 1 без пятки
    Dim i As Integer
    Dim Index As String
    
    If PaidMod < 0 Then
        ClearSheet WP, Range("HDR_WP")
        FrL = EOL(WP) + 1                       ' начальная строка Платежей на WP
        P_Lines = AutoFilterReset(1) - 3
        Range("A1:AC" & EOL(1) - 3).AutoFilter Field:=9, _
            Criteria1:="=" & Range("I" & Prow)  ' фильтр в 1 по выбранной Организации
        Give "A2:AC" & P_Lines, "A2", 1  ' все Платежи Организации в WP
        PaidMod = -PaidMod
        ToL = EOL(WP)
    End If
       
    Select Case PaidMod
        Case 1                      ' режим 1 - выводим только выбранную строку
            HideLns FrL, ToL, 0                 ' Hide все
            Index = Sheets(1).Cells(Prow, 2)
            HideLns FrL, ToL, 2, Index, False   ' Unhide выбранный Платеж
        Case 2                      ' режим 2 - выводим только Платежи, которых нет в SF
            HideLns FrL, ToL, 0                 ' Hide все
            HideLns FrL, ToL, 4, , False        ' Unhide Платежи, которых нет в SF
        Case 3                      ' режим 3 - все Платежи выбранной Организации
            HideLns FrL, ToL, 0, , False        ' Unhide все
        Case Else
            MsgBox "ERROR in WPaid: PaidMod = " & PaidMod, , "ERROR!"
    End Select
'    If Prow <= 4 Or Prow > Lines Then Prow = Lines
'
'    Prow = ActiveCell.Row                               ' указана ячейка (Prow,Pcol)
'    Pcol = ActiveCell.Column
'    If Prow > 4 And Prow <= Lines And Pcol < 17 Then    ' если она в допустимом диапазоне
'        Range(Cells(Prow, 6), Cells(Prow, 17)).Select
'        Selection.Interior.Color = rgbLightCoral        ' закрасим строку
'        Range(Cells(Prow, Pcol)).Interior.Color = rgbCoral
'        Selection.Interior.Color = rgbCoral             '   .. и более ярко - ее саму
'    Else
'        Prow = Lines: Pcol = 6
'    End If
End Sub
Sub WP_Paid_Click()
'
' Переключение статусной переменной по выводу Платежей на рабочем ласте
'   19.1.12
    
    PaidMod = PaidMod + 1
    If PaidMod = 3 Then PaidMod = 0
    Call WPaid
'
'    Call Wopp
'    Call Wcontr     ' вывод Договоров в режиме ContrFlag
'    Call WADSK      ' вывод имеющихся в Организации продуктов Autodesk в режиме ADSKFlag

End Sub
Sub Wopp(Account)
'
' Вывод списка Проектов на листе WP для выбранной Организации (Account)
'   25/1/2012
'   10.2.2012 - ревизия

    Dim OppLines, WPlines, i As Integer
    
    Lines = EOL(WP) + 2                     ' начальная строка для Проектов на WP
    AutoFilterReset SFopp
    OppLines = EOL(SFopp) - SFresLines      ' всего строк в SFopp
    Range("A1:D" & OppLines).AutoFilter Field:=4, _
        Criteria1:="=" & Account            ' фильтр SFopp по выбранной Организации
   
    Give "B1:B" & OppLines, "B" & Lines, 2  ' Номер Проекта
    WPlines = EOL(WP)                       ' число строк WP вместе с Проектам
    ActiveSheet.Range("B1:B" & WPlines).HorizontalAlignment = xlCenter
    Give "C1:C" & OppLines, "K" & Lines, 2  ' Имя Проекта
    Give "H1:I" & OppLines, "G" & Lines, 2  ' Даты создания и закрытия Проекта
    Give "J1:J" & OppLines, "L" & Lines, 2  ' Продавец
    Give "O1:O" & OppLines, "J" & Lines, 2  ' Еще не доплачено (руб/USD/EUR)
    Give "P1:P" & OppLines, "P" & Lines, 2  ' Не оплачено %
    Give "M1:M" & OppLines, "N" & Lines, 2  ' Стоимость
    Give "G1:G" & OppLines, "D" & Lines, 2  ' Флаг "Реализовано"
    Give "E1:E" & OppLines, "F" & Lines, 2  ' Вероятность (%)
    Give "Q1:Q" & OppLines, "I" & Lines, 2  ' Тема&Вид деятельности
    Give "N1:N" & OppLines, "A" & Lines, 2  ' Валюта Проекта

'**** Проход по извлеченным Проектам
    For i = Lines + 1 To WPlines        ' Подправляем формат Проектов на WP
        PerCent i, 6                    ' Вероятность %
        PerCent i, 16                   ' Не оплачено %
        CurCode i, 10, 1                ' формат валюты Стоимости
        CurCode i, 14, 1                ' формат валюты "Еще не доплачено"
'        If Cells(i,6) = "100%  ' окрашиваем оплаченный Проект
    Next i
    
    Range("F1").Select                  ' копируем формат заголовка из Платежей
    Selection.Copy
    Range("B" & Lines & ":Q" & Lines).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Rows(Lines & ":" & WPlines).Font.Name = "Calibri"
    Rows(Lines & ":" & WPlines).Font.size = 8
    Range("A" & Lines + 1 & ":A" & WPlines).WrapText = False ' поджимаем первую колонку
    Range("A" & Lines & "," & "D" & Lines).Clear
    
    HideLns Lines, WPlines, 6, "100%"

End Sub
Sub Wcontr(Account)
' Вывод списка Договоров на листе WP для выбранной Организации (Account'a)
'   25/1/2012
' 10.2.2012 - ревизия
   
    Dim WPlines, ContrLines, Index, i
   
    Lines = EOL(WP) + 2                     ' начальная строка для Договоров на WP
    AutoFilterReset SFD                       ' сбрасываем все фильтры в отчетах на 3
'    SheetSort 3, 4                          ' сортируем Договоры по порядку в 1С
    ContrLines = EOL(SFD) - SFresLines      ' всего строк в SFD
    Range("A1:F" & ContrLines).AutoFilter Field:=6, _
        Criteria1:="=" & Account            ' фильтр SFD по выбранной Организации
    
    Give "A2:A" & ContrLines, "B" & Lines + 1, 3   ' Индекс Договора (ссылка на лист 3)
    WPlines = EOL(WP)                       ' число строк WP вместе с Договорами
    ActiveSheet.Range("B1:B" & WPlines).HorizontalAlignment = xlCenter
    Give "M1:M" & ContrLines, "F" & Lines, 3   ' Статус Договора
    Give "D1:E" & ContrLines, "M" & Lines, 3  ' Договор и Основной Договор
    Give "G1:G" & ContrLines, "I" & Lines, 3  ' Вид Договора
    Give "H1:I" & ContrLines, "G" & Lines, 3  ' Даты начала и завершения Договора
    Give "P1:P" & ContrLines, "K" & Lines, 3  ' Проект
    Give "Q1:Q" & ContrLines, "P" & Lines, 3  ' Номер Проекта

'**** Проход по извлеченным Договорам
    Sheets(WP).Select
    For i = Lines + 1 To WPlines
        Index = Cells(i, 2)     ' Индекс Проекта - номер на листе Договоров из 1С
        Sheets(3).Range("E" & Index + 1).Copy
        Cells(i, 1).Select
        ActiveSheet.Paste
'        PerCent i, 6                    ' Вероятность %
'        PerCent i, 16                   ' Не оплачено %
'        CurCode i, 10, 1                ' формат валюты Стоимости
'        CurCode i, 14, 1                ' формат валюты "Еще не доплачено"
'        If Cells(i,6) = "100%  ' окрашиваем оплаченный Проект
    Next i
    
    Range("F1").Select                  ' копируем формат заголовка из Платежей
    Selection.Copy
    Range("B" & Lines & ":Q" & Lines).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    Rows(Lines & ":" & WPlines).Font.Name = "Calibri"    ' мелкий шрифт
    Rows(Lines & ":" & WPlines).Font.size = 8
    Range("A" & Lines & ":A" & WPlines).WrapText = False
    Range("A" & Lines & "," & "D" & Lines).Clear    ' стираем длинные заголовки колонок
    
    HideLns Lines, WPlines, 6, "Закрыт"
    HideLns Lines, WPlines, 6, "Нет в SF"

End Sub
Sub WADSK(Account)
' Вывод списка продуктов Autodesk на листе WP для выбранной Организации
'   12.2.2012

End Sub

Sub AccAddFrom1C(Acc As String)
'
' Попытка разбора и добавление Организации Acc из 1С в SF
'   5.2.2012

'    Dim Row1C   ' индекс - номер строки Организации в 1С
    Dim adr As String
    Dim AdrLength As Integer
    Dim AdrArr() As String
    
    Dim PostIndex As Integer
    Dim City, Street, Region, Country As String
'    Dim Q
    
    
    ModStart Acc1C, "поиск Организации " & Acc & " из 1C"
    
    adr = Adr1C(Acc)
    adr = Trim(adr)
    AdrLength = Len(adr)
    
    AddressForm.TextBox1 = Acc
    AddressForm.TextBox2 = adr
    
    AdrArr = split(adr, ",")
    AddressForm.PostIndex = AdrArr(0)
    AddressForm.City = AdrArr(1)
    AddressForm.Street = AdrArr(2)
    AddressForm.PostCountry = "Россия"
    
    AddressForm.Show
    
  '  PostIndex = Left(Adr, 6)
    MsgBox "Адрес в справочнике 1С=" & adr & " Индекс=" & PostIndex
        
    NewSFaccForm.Show
'    Do
'        AccSF = SearchSFacc(Acc)
'        If AccSF <> 0 Then
'            Q = MsgBox("Это соответствует " & AccSF & " в SF?", vbYesNoCancel)
'            If Q = vbYes Or Q = vbCancel Then
'                Exit Do
'            End If
'        End If
'    ' модификация строки - имени Асс
''?????????????????????????????????????????????
'    Loop
        
    ModEnd WP
End Sub
Sub SearchSFacc(txt)
'
' поиск в списке Организаций SFacc подстроки txt по имени Организаций
' возвращает номер строки в SFacc или 0, если не найдено
'   3.2.2012

    Dim Row, i As Integer
    Dim S, Str As String

  
    For Row = 2 To Lines
        Str = Sheets(SFacc).Cells(Row, 1)
S = Application.WorksheetFunction.Search(txt, Str)
'        i = Search(txt, Str)
        MsgBox i
    Next Row
End Sub
Sub Give(Fr1, ToWP, Mode)
'
' подпрограмма копирования ячейек Fr1 на место ToWP на рабочий лист WP
' если  Mode = 1 - копирование PasteSpecial из Платежей
'       Mode = 2 - копируем из SFopp
'       Mode = 3 - копируем из SFD
' 22.1.2012

    Select Case Mode
        Case 1
            Sheets(1).Range(Fr1).Copy
            Sheets(WP).Select
            Range(ToWP).Select
            Selection.PasteSpecial Paste:=xlPasteColumnWidths, _
                Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Case 2
            Sheets("SFopp").Range(Fr1).Copy
            Sheets(WP).Select
            Range(ToWP).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
        Case 3
            Sheets("SFD").Range(Fr1).Copy
            Sheets(WP).Select
            Range(ToWP).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
    End Select

    ActiveSheet.Paste
End Sub
