Attribute VB_Name = "From1C"
'---------------------------------------------------------------------------
' Макросы для загрузки отчетов из 1С
'
' From1Cpayment  - заменяет лист отчета из 1С "Приход денег на счета"
' From1Cdogovor  - заменяет лист отчета 1С "Договоры" новым из 1С
' From1Caccount  - заменяет лист отчета 1С "Клиенты .." новым из 1С
' FromStock      - замена Складской Книги в листе Stock
' П.Л.Храпкин 30.4.2012
'   5.2.2012 - общая опработка штампов и пяток
'   7.2.2012 - входной отчет хорош?
'   29.2.2012 - bug fix, ревизия
'   6.3.12 - Options Explicit, доопределение констант и переменных
'  22.4.12 - дописан From1Caccount
'  17.5.12 - загрузка Складской книги

Option Explicit
Sub From1Cpayment()
'
' Заменяет лист отчета 1С "Приход денег на счета" новым Листом на первом месте
'   15.1.2012
'   10.2.2012 - проверка штампов входного и прежнего отчетов
'   29.2.2012 - ревизия, EOL
'   6.3.12 - Option Explicit
'  30.4.12 - SFaccColFill

    Dim LO, Ln, PaymentsheetName, Pold

    ModStart 2, "Заносим новый отчет 1С по платежам в Match SF-1C"

    CheckSheet 1, 1, 1, Stamp1Cpay1    ' проверяем правильность нового отчета
    CheckSheet 1, 1, 2, Stamp1Cpay2
    CheckSheet 2, 1, 6, Stamp1Cpay1    ' прежний отчет на месте?
    CheckSheet 2, 1, 7, Stamp1Cpay2
    
    Sheets(1).Select                   ' новый отчет в листе 1
    Sheets(1).UsedRange.Activate

    With Selection.Font    ' устанавливаем по новому отчету стандартный штифт
        .Name = "Calibri"
        .size = 8
    End With
    
    Sheets(2).Columns("A:E").Copy   ' из прежнего отчета копируем колонки A:E
    Sheets(1).Columns("A:A").Select '    и вставляем их слева к новому отчету
    Selection.Insert Shift:=xlToRight
        
    Range("J:Q,T:U,W:X").Select           ' делаем невидимыми ненужные колонки
    Selection.EntireColumn.Hidden = True  '   ..валютных проводок, расходных
                                          '   .. кредитов,отделов и фирм
    LinesOld = EOL(2)               ' кол-во строк в старом отчете
    Lines = EOL(1)                  ' кол-во строк в новом отчете

' дополняем колонки формул до конца рабочей области
    LO = LinesOld - 3
    Ln = Lines - 2
    Range(Cells(LO, 1), Cells(LO, 5)).Select
    If Ln > LO Then
        Selection.AutoFill Destination:=Range(Cells(LO, 1), Cells(Ln, 5)), _
            Type:=xlFillDefault
    End If
    
    Sheets(2).Select                ' вставляем группу ячеек - итоги сверки
    Range(Cells(LinesOld - 2, 2), Cells(LinesOld, 18)).Copy
    Sheets(1).Select
    Range(Cells(Lines - 1, 2), Cells(Lines - 1, 2)).Activate
    ActiveSheet.Paste
      
    SFaccColFill PAY_SHEET  ' в колонке 1 если Организация есть в SF
    SFaccCol PAY_SHEET, PAY_RESLINES    ' раскрашиваем колонку A
  
'*******************************************************
    Call PaymentPaint    '* раскрашиваем Лист Платежей *
'*******************************************************
    
' подставляем название нового отчета в лист SF
    PaymentsheetName = Sheets(1).Name   ' имена листов старого и нового
    Pold = Sheets(2).Name               '   ..отчетов по Платежам
    
    SheetsCtrlH "SF", Pold, PaymentsheetName
    SheetsCtrlH "P_PaidContract", Pold, PaymentsheetName
'    SheetsCtrlH "P_Update", Pold, PaymentsheetName
        
    Sheets(2).Select
    ActiveWindow.SelectedSheets.Delete  ' удаляем старый платежный отчет
    
    Sheets(PaymentsheetName).Select
    With ActiveWorkbook.Sheets(1).Tab   ' Таб нового отчета - красный
        .Color = 255
        .TintAndShade = 0
    End With
    
    Doing = "Отчет платежей " & Sheets(1).Name
    ModEnd 1
End Sub
Sub From1Cdogovor()
'
' Заменяет лист отчета 1С "Договоры" новым Листом на первом месте
'
' 2.1.2012 П.Л.Храпкин
' 6.1.2012 внедрение колонки Н - <Основной>/<Договор>
' 12.1.2012 переписано с использованием подпрограмм
' 20.1.2012 не работает - отладка и ревизия
' 6.3.12 - константы по пятке и штамп
' 30.4.12 - SFaccColFill
' 17.4.12 - bug fix с новым отчетом меньше старого
    
    Dim Total, BadContr

    LinesOld = ModStart(4, "Обновляем из 1С отчет по Договорам") - DOGRES
    
    CheckSheet 4, 1, 10, Stamp1Cdog1    ' штамп Договора
    
    Lines = EOL(1) - 2      ' кол-во строк - Договоров в новом отчете
    
    Sheets(4).Columns("A:H").Copy   ' в прежнем отчете копируем колонки A:H
    Sheets(1).Select
    Sheets(1).Columns("A:A").Select '     .. и вставляем их слева нового отчета
    Selection.Insert Shift:=xlToRight
                                    ' если в новом отчете рекордов больше, то
    If Lines > LinesOld Then        '     .. дополняем колонки формул до конца
        Range(Cells(LinesOld, 1), Cells(LinesOld, 8)).Select
        Selection.AutoFill Destination:= _
            Range(Cells(LinesOld, 1), Cells(Lines, 8)), Type:=xlFillDefault
    Else
        Range(Cells(Lines + 1, 1), Cells(LinesOld, 8)).Select
        Selection.Delete Shift:=xlUp
    End If
    
    Sheets(4).Range("A" & LinesOld + 1 & ":O" & LinesOld + DOGRES).Copy
    Sheets(1).Select                ' вставляем группу ячеек - итоги сверки
    Range("A" & Lines + 1).Select
    ActiveSheet.Paste               ' вставляем группу ячеек - итоги сверки
    
    Sheets(1).Move After:=Sheets("SF") ' перемещаем новый отчет после листа SF
    Sheets(3).UsedRange.Activate
    Rows.RowHeight = 15             ' высота строк до конца = 15
    
    SheetsCtrlH "SFD", Sheets(4).Name, Sheets(3).Name ' обновляем формулы в SFD
' >>>>>>>>>>>>>>>> здесь Ctrl/H для других листов Match

    Total = EOL("SFD")
    BadContr = Cells(Total, 1)
    
    MsgBox "В SF " & Total - 7 & " Договоров, из них " & BadContr _
                                                & " не соответствует 1С"
    Sheets(4).Delete                ' удаляем старый отчет по Договорам
    
    SFaccColFill DOG_SHEET  ' в колонке 1 если Организация есть в SF
    SFaccCol DOG_SHEET, DOGRES    ' раскрашиваем колонку A
      
'************************************************************
    Call ContractPaint       '* раскрашиваем лист Договоров *
'************************************************************

    ActiveWorkbook.Sheets(3).Tab.Color = LimeG  ' Tab нового отчета - зеленый
    Doing = Sheets(3).Name
    Call ModEnd(3)
End Sub
Sub From1Caccount()
'
' Заменяет лист отчета 1С "Список клиентов 1С" новым Листом на первом месте
'
' 22.4.2012 П.Л.Храпкин
' 30.4.12 - SFaccColFill

    ModStart Acc1C, "Обновление листа Справочника клиентов 1С"

    LinesOld = EOL(Acc1C) - ACC1C_RES   ' кол-во строк в старом отчете
    Lines = EOL(1) - ACC1C_RES          ' кол-во строк в новом отчете
    
    CheckSheet 1, ACC1C_RES + 1, 2, ACC1C_STAMP
    CheckSheet Acc1C, 1, A1C_NAME_COL, ACC1C_STAMP
    
    ' переновим верхние строки в пятку
    Sheets(1).Select
    Rows("1:" & ACC1C_RES).Cut
    Cells(Lines + ACC1C_RES + 2, 1).Select
    ActiveSheet.Paste
    Rows("1:" & ACC1C_RES).Delete
    
    Sheets(Acc1C).Columns("A:C").Copy   ' из прежнего отчета копируем колонки A:E
    Sheets(1).Columns("A:A").Select     '    и вставляем их слева к новому отчету
    Selection.Insert Shift:=xlToRight
    
' дополняем колонки формул до конца рабочей области
    If LinesOld < Lines Then
        Range(Cells(LinesOld, 1), Cells(LinesOld, 3)).Select
        Selection.AutoFill Destination:=Range(Cells(LinesOld, 1), Cells(Lines, 3)), _
            Type:=xlFillDefault
    Else
        Range(Cells(Lines + 1, 1), Cells(LinesOld + ACC1C_RES, 3)).CleaContents
    End If
    
    Sheets(Acc1C).Name = "RRR"          ' прежний отчет переименовываем в RRR
    Worksheets(1).Name = Acc1C          ' новый отчет переименовываем в SF
    
    Sheets(Acc1C).UsedRange.Activate
    Rows.RowHeight = 15                 ' высота строк до конца = 15
    
    Call SheetsCtrlH(2, "RRR!", Acc1C)  ' заменяем ссылки в формулах Платежей 1С
    Call SheetsCtrlH(4, "RRR!", Acc1C)  '   и в Договорах
    
    Sheets(Acc1C).Move After:=Sheets("RRR") ' перемещаем новый отчет после листа Acc1C
    Sheets("RRR").Delete
    
    SFaccColFill Acc1C              ' в колонке 1 если Организация есть в SF
    SFaccCol Acc1C, ACC1C_RES       ' раскрашиваем колонку A
      
    Sheets(Acc1C).Tab.Color = rgbRed    ' окрашиваем Tab нового отчета в красный цвет
    ModEnd Acc1C
End Sub
Sub FromStock()
'
' (*) FromStock() - замена Складской Книги в листе Stock
'       17/5/12
'   18.5.12 - добавил колонку ContrADSK


    Dim LO As Integer, Ln As Integer    ' кол-ва строк в старом и новом отчетах
    Dim MSG As String
    
    LO = ModStart(STOCK_SHEET, "Обновление Складской Книги")
    Ln = EOL(1)

    CheckSheet 1, 3, STOCK_PRODUCT_COL - STOCK_MYCOLS, STOCK_STAMP
    CheckSheet STOCK_SHEET, 1, STOCK_PRODUCT_COL, STOCK_STAMP
    
    Sheets(1).Select
    Rows("1:2").Select        ' удаляем первые две строки
    Selection.Delete Shift:=xlUp
    
    Sheets(STOCK_SHEET).Columns("A:F").Copy ' из прежнего отчета копируем колонки A:F
    Sheets(1).Columns("A:A").Select         '    и вставляем их слева к новому отчету
    Selection.Insert Shift:=xlToRight
                                                                
    Sheets(STOCK_SHEET).Name = "Tmp"  ' переименовываем прежний отчет в "Tmp"
    Sheets(1).Name = STOCK_SHEET      '   .. а новый в "Склад"

    Sheets(STOCK_SHEET).Move Before:=Sheets("Tmp")  ' перемещаем новый отчет
    Sheets("Tmp").Delete                ' уничтожаем прежний отчет SF
    Sheets(STOCK_SHEET).Tab.Color = rgbBlue ' окрашиваем Tab нового отчета
    
    MSG = "В прежней Складской Книге " & LO & " строк, в новой "
    If LO = Ln Then MSG = MSG & "тоже "
    MsgBox MSG & Ln
        
    ModEnd STOCK_SHEET
End Sub
