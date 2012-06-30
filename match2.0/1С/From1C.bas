Attribute VB_Name = "From1C"
'---------------------------------------------------------------------------
' Макросы для загрузки отчетов из 1С
'
' From1Cpayment  - заменяет лист отчета из 1С "Приход денег на счета"
' From1Cdogovor  - заменяет лист отчета 1С "Договоры" новым из 1С
' From1Caccount  - заменяет лист отчета 1С "Клиенты .." новым из 1С
' FromStock      - замена Складской Книги в листе Stock
'
' 29.6.2012 П.Л.Храпкин match 2.0

Option Explicit
Sub From1Cpayment()
'
' From1Cpayment() - Заменяет лист отчета 1С "Приход денег на счета" новым Листом на первом месте
'
'  25.6.12 - match 2.0
'  27.6.12 - избавляемся от формул со ссылками на другой файл

    Dim LO, Ln, Lines, LinesOld, PaymentsheetName, Pold

    ModStart REP_1C_P_LOAD
    
    Sheets(1).Select                   ' новый отчет в листе 1
    Sheets(1).UsedRange.Activate
    With Selection.Font    ' устанавливаем по новому отчету стандартный шрифт
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
    
    Call DateCol(PAY_SHEET, PAYDATE_COL) ' преобразование колонки Дат
    SheetSort PAY_SHEET, PAYDATE_COL     ' сортируем Платежи по Дате прихода денег
        
    Sheets(2).Select                ' вставляем группу ячеек - итоги сверки
    Range(Cells(LinesOld - 2, 2), Cells(LinesOld, 18)).Copy
    Sheets(1).Select
    Range(Cells(Lines - 1, 2), Cells(Lines - 1, 2)).Activate
    ActiveSheet.Paste
    
    SFmatchFill PAY_SHEET   '*** выполняем Update по отчетам SF
  
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
Sub SFmatchFill(SheetN)
'
' - SFmatchFill SheetN
'       заполнение колонок PAYINSF_COL и PAYSFLN_COL листа PAY_SHEET
'       в зависимости от наличия значения в колонке SF_COD_COL листа SF
' 30.6.12

    Dim L As Integer        '= EOL заполняемого отчета
    Dim iPaid As Integer    '= номер "сшитой" строки в SF
    Dim AccCol As Integer   '= номер "сшиваемой" колонки
    Dim i As Integer        '= номер текущей строки заполняемого отчета
    Dim j As Integer        '= номер строки SF_PA
    
'    SFaccColFill PAY_SHEET  ' в колонке 1 если Организация есть в SF
'    SFaccCol PAY_SHEET, PAY_RESLINES    ' раскрашиваем колонку A
        
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
            ErrMsg FATAL_ERR, "неправильный лист в SFaccColFill"
            Stop
    End Select
        
        
    Set DB_SFDC = Workbooks.Open(F_SFDC, UpdateLinks:=False, ReadOnly:=True)
    With ThisWorkbook.Sheets(PAY_SHEET)
        For i = 2 To L
            Progress i / L
                '-- "нормализуем" имя Организации  и "сшиваем" его с SFacc --
            .Cells(i, AccCol) = Replace(Compressor(.Cells(i, AccCol)), vbCrLf, "")
            
            If IsMatch(.Cells(i, AccCol), SFACC_ACC1C_COL, , SFacc, DB_SFDC) Then
                .Cells(i, PAYISACC_COL) = "1"
            Else
                .Cells(i, PAYISACC_COL) = ""
            End If
                        
                '-- "сшиваем" с Платежом в SF --
            If IsMatch(.Cells(i, PAYCODE_COL), SF_COD_COL, iPaid, SF, DB_SFDC) Then
                .Cells(i, PAYINSF_COL) = "1"
                .Cells(i, PAYSFLN_COL) = iPaid
                SFid = DB_SFDC.Sheets(SF).Cells(iPaid, SF_PAYID_COL)
            Else
                .Cells(i, PAYINSF_COL) = ""
                .Cells(i, PAYSFLN_COL) = ""
                SFid = ""
            End If
            
                '-- "сшиваем" с Контрактами ADSK в SF_PA --
    '..... потом здесь можно поместить ВСЕ Контракты ADSK по номерам с "+"
            If IsMatch(SFid, SFPA_PAYID_COL, , SF_PA, DB_SFDC) Then
                .Cells(i, PAYADSK_COL) = "1"
            Else
                .Cells(i, PAYADSK_COL) = ""
            End If
        Next i
    End With
    DB_SFDC.Close SaveChanges:=False
End Sub
Sub testSFmatchFill()
    EOL_PaySheet = EOL(PAY_SHEET) - PAY_RESLINES
    SFmatchFill PAY_SHEET
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
' 16.6.12 - ревизия, убрал перенос колонок-формул из старого отчета

    ModStart Acc1C, "Обновление листа Справочника клиентов 1С"

    LinesOld = EOL(Acc1C) - ACC1C_RES   ' кол-во строк в старом отчете
    Lines = EOL(1) - ACC1C_RES          ' кол-во строк в новом отчете
    
    CheckSheet 1, 4, 2, ACC1C_STAMP
    CheckSheet Acc1C, 1, A1C_NAME_COL, ACC1C_STAMP
    
    ' переноcим верхние строки в пятку
    Sheets(1).Select
    Rows("1:" & ACC1C_RES).Cut
    Cells(Lines + ACC1C_RES + 2, 1).Select
    ActiveSheet.Paste
    Rows("1:" & ACC1C_RES).Delete
    
    ' вставляем колонку 1
    Sheets(1).Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
        
    Sheets(Acc1C).Name = "RRR"          ' прежний отчет переименовываем в RRR
    Worksheets(1).Name = Acc1C          ' новый отчет переименовываем в "Список клиентов 1С"
    
    Sheets(Acc1C).UsedRange.Activate
    Rows.RowHeight = 15                 ' высота строк до конца = 15
    
    Sheets(Acc1C).Move After:=Sheets("RRR") ' перемещаем новый отчет после листа Acc1C
    Sheets("RRR").Delete
    
    SFaccColFill Acc1C              ' заполняем в колонке 1 если Организация есть в SF
    SFaccCol Acc1C, ACC1C_RES       ' раскрашиваем колонку: желтый - есть в SF, красный - нет
      
    Sheets(Acc1C).Tab.Color = rgbRed    ' окрашиваем Tab нового отчета в красный цвет
    ModEnd Acc1C
End Sub
Sub FromStock()
'
' (*) FromStock() - замена Складской Книги в листе Stock
'       17/5/12
'   18.5.12 - добавил колонку ContrADSK


    Dim LO As Integer, Ln As Integer    ' кол-ва строк в старом и новом отчетах
    Dim Msg As String
    
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
    
    Msg = "В прежней Складской Книге " & LO & " строк, в новой "
    If LO = Ln Then Msg = Msg & "тоже "
    MsgBox Msg & Ln
        
    ModEnd STOCK_SHEET
End Sub
