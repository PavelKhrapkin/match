Attribute VB_Name = "FromSF"
'---------------------------------------------------------------------------
' Обновление отчетов из salesforce.com в Match SF-1C.xlms
'
' Подпрограммы и Макросы (*):
' (*) ShowDBGbanel - вывод панели для отладки сегментов по Ctrl/Shift/Q
' <*> Match1C-SF()- обновляет лист SF- отчетов по Платежам для сверки SF с 1С
' <*> SFDreport() - заменяет лист SFD новым отчетом по Договорам из Salesforce
' <*> SFaccRep()  - заменяет лист SFacc новым отчетом по Организациям из SF
'
' П.Л.Храпкин 2.1.2012
'   5.2.2012 - работы по параметризации имен и позиций листов, штампов
'   8.2.2012 - проверка (CheckSheet) "Вставляемый отчет хорош?"
'  29.2.2012 - Не копируем шапку в новый отчет SFopp
'   4.3.12 - bug fix
'  30.3.12 - bug fix EOL в SFD
'  28.4.12 - добавлен лист SForders
'  30.4.12 - SFaccColFill
'  12.5.12 - bug fix
'  16.5.12 - новый отчет по свзкам Платежей с Контрактами ADSK SF_PA
'  14.7.12 - match2.0 - все неспецифические действия выполняет MoveToMatch

    Option Explicit    ' Force explicit variable declaration
        
Sub ShowControlPanel()
Attribute ShowControlPanel.VB_Description = "8.2.2012 Запуск ShowDBGpanel "
Attribute ShowControlPanel.VB_ProcData.VB_Invoke_Func = "Q\n14"
'
' Вывод отладочной управляющей панели с командными кнопками по всем отдельным листам
'
' Ctrl/Shift/Q
'
'   8.2.2012

    MainControlPanel.Show
    End
End Sub
Sub Match1C_SF()
'
' (*) обновление отчета SF "Платежи Сверка с 1С"
'
' запускается кнопками [SF] на листе Платежей и "SF"
'
'   4.9.11
'   5.1.12 - использование подпрограммы PaintPayment, Application.StatusBar
'   9.1.12 - корректное копирование сводки по SF
'  26.1.12 - проверка, что на входе действительно отчет Платежи, сортировка SF
'  28.1.12 - параметризация по именам листов
'  14.8.12 - match2.0

    ModStart REP_SF_LOAD
'''    CheckSheet SF
    EOL_SF = EOL(SF) - SFresLines
    
    InsMyCol "SF_MyCol", EOL_SF
    InsSummary "SF_Summary", EOL_SF + SFresLines

'---- заполняем Match - связку с Платежами 1С
    Set DB_1C = Workbooks.Open(F_1C, UpdateLinks:=False, ReadOnly:=True)
    DB_1C.Sheets(PAY_SHEET).Select

    Dim i
    With ThisWorkbook.Sheets(SF)
        For i = 2 To EOL_SF
            Progress i / EOL_SF
            .Cells(i, 1) = CSmatch(.Cells(i, SF_COD_COL), PAYCODE_COL)
        Next i
    End With
'********************
'    PaymentPaint   '*
'********************

    DB_1C.Close
    ModEnd
 End Sub
 Sub InsMyCol(F, EL)
 '
 ' - InsMyCol(SheetN, F) - вставляем колонки формул F в SheetN
 '  14.7.12
 
    Dim i As Integer
    For i = 1 To Range(F).Columns.count
        ActiveSheet.Cells(1, 1).EntireColumn.Insert
        ActiveSheet.Columns(i).ColumnWidth = Range(F).Cells(3, i)
    Next i
    Sheets("Forms").Range(F).Copy Destination:=ActiveSheet.Cells(1, 1)
    ActiveSheet.Range("A2:B" & EL).FillDown
End Sub
 Sub InsSummary(F, EL)
 '
 ' - InsSummary(SheetN, F) - вставляем сводку F в ряд EL от конца вверх
 '  14.7.12
 
    Sheets("Forms").Range(F).Copy _
        Destination:=ActiveSheet.Cells(EL - Range(F).Rows.count + 1, 1)
 End Sub
'''''''''''    LinesOld = ModStart("SF", _
'''''''''''        "MatchSF_1C - обновляем лист SF по отчету Salesforce <Платежи из 1С>")
'''''''''''
'''''''''''    Lines = Sheets(1).UsedRange.Rows.count      ' кол-во строк в новом отчете
'''''''''''    CheckSheet 1, Lines - 4, 1, SFpayRepName    ' проверяем правильность нового и
'''''''''''    CheckSheet "SF", LinesOld + 2, 3, SFpayRepName  '   .. прежнего отчетов SF
'''''''''''    SheetSort 1, 3                      ' новый отчет сортируем по датам
'''''''''''
'''''''''''    Sheets("SF").Select
'''''''''''    Columns("A:B").Copy                 ' копируем 2 правые колонки
'''''''''''
'''''''''''    Sheets("SF").Name = "RRR"           ' прежний отчет SF переименовываем в RRR
'''''''''''
'''''''''''    Worksheets(1).Name = "SF"           ' новый отчет переименовываем в SF
'''''''''''    Sheets("SF").Select
'''''''''''    Columns("A:A").Select
'''''''''''    Selection.Insert Shift:=xlToRight   ' вставляем 2 колонки из прежнего листа
'''''''''''
'''''''''''    Rows("2:" & Lines).RowHeight = 15
'''''''''''
'''''''''''    Sheets("RRR").Range("E" & LinesOld - 1 & ":G" & LinesOld).Copy
'''''''''''    Sheets("SF").Range("E" & Lines - 1).Select
'''''''''''    ActiveSheet.Paste                   ' копируем оборот CSIT
'''''''''''    Sheets("RRR").Range("A" & LinesOld - 1 & ":B" & LinesOld).Copy
'''''''''''    Sheets("SF").Range("A" & Lines - 1).Select
'''''''''''    ActiveSheet.Paste   ' копируем красную сводку - "не сошлось с Платежами 1С"
'''''''''''
'''''''''''' если надо, продлеваем вниз колонки формул
'''''''''''    If Lines > LinesOld Then Range("A" & LinesOld - SFresLines & _
'''''''''''                                  ":B" & Lines - SFresLines).FillDown
'''''''''''
'''''''''''    With Range("A" & Lines - 5 & ":B" & Lines - 2)
'''''''''''        .ClearContents
'''''''''''        .Interior.Color = rgbWhite      ' очищаем под колонкой до блока итогов
'''''''''''    End With
'''''''''''
'''''''''''    Call SheetsCtrlH(2, "RRR!", "SF!")  ' заменяем ссылки в формулах Платежей 1С
''''''''''''    Call SheetsCtrlH("P_Update", "RRR!", "SF!") '   в DL Payment_Update
'''''''''''
'''''''''''    Sheets("SF").Move After:=Sheets(2)  ' переносим новый отчет SF на второе место
'''''''''''    Sheets("RRR").Select
'''''''''''    ActiveWindow.SelectedSheets.Delete  ' уничтожаем старый отчет SF
'''''''''''
''''''''''''********************
'''''''''''    PaymentPaint   '*
''''''''''''********************
'''''''''''
'''''''''''    Sheets("SF").Tab.Color = rgbViolet ' окрашиваем Tab нового отчета SF
'''''''''''    ModEnd 1

Sub SFDreport()
'
' 3. Заменяет лист SFD новым отчетом по договорам из SF
'
'   6.1.2012 - при замене SFD CTRL/H в P_ листах; колонки A и В - ключ и проверка SFD
'  26.1.2012 - проверка правильности листа CheckSheet
'   4.3.12 - bug fix FillDown
'  30.3.12 - bux fix -- EOL вместо RowsCount
    
    Dim LO, Ln As Integer           ' количества строк в старом и новом отчетах
    Dim OOO, NNN
    
    ModStart "SFD", "Обновление листа SFD - отчета Salesforce по Договорам <SFD>"

    LinesOld = EOL(SFD)       ' кол-во строк в старом отчете
    Lines = EOL(1)            ' кол-во строк в новом отчете
    LO = LinesOld - SFresLines
    Ln = Lines - SFresLines
    
    CheckSheet 1, Lines - 4, 1, SFcontrRepName
    CheckSheet "SFD", LinesOld - 4, 3, SFcontrRepName
    
    AutoFilterReset 4           ' сбрасываем фильтры в отчете 1С по договорам
    AutoFilterReset "SFD"       '     .. и в листе SFD
    SheetSort 4, 4              ' сортируем Договоры по порядку номеров в 1С
    
    Sheets("SFD").Select
    Sheets("SFD").Name = "RRR"                      'переименовываем SFD в RRR
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst

    Sheets(1).Name = "SFD"                          'переименовываем новый отчет в SFD
    Sheets("SFD").Select
    Sheets("SFD").Move Before:=Sheets(5)            'и перемещаем его перед RRR

    Sheets("RRR").Columns("A:B").Copy               ' копируем 2 правые колонки
    Sheets("SFD").Select
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight       ' вставляем 2 колонки из прежнего

    Rows("2:" & Lines).Select
    Selection.RowHeight = 15
    
    Worksheets("RRR").Range("A" & LinesOld).Copy
    Worksheets("SFD").Range("A" & Lines).Select
    ActiveSheet.Paste       ' копируем красную сводку - не совпавшие Договоры
    
    OOO = Worksheets("RRR").Range("A" & LinesOld)
    NNN = Worksheets("SFD").Range("A" & Lines)
    
    MsgBox "Найдено " & NNN & " не совпавших Договоров, было " & OOO
    
    If Ln > LO Then
        Range("A" & LO - 1 & ":B" & Ln).Select
        Selection.FillDown      ' дополняем колонки формул в новом отчете до конца
    End If
       
    With Range("A" & Lines - 5 & ":B" & Lines - 1)
        .ClearContents
        .Interior.Color = rgbWhite  ' очищаем под колонкой до блока итогов
    End With

    SheetsCtrlH 3, "RRR!", "SFD!"   ' замена ссылок в формулах отчетов на SFD
    SheetsCtrlH "P_PaidContract", "RRR!", "SFD!"
    
    Sheets("RRR").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("SFD").Select
    ActiveWorkbook.Sheets("SFD").Tab.Color = rgbViolet

'********************
    ContractPaint  '*
'********************
    
    ModEnd 3
End Sub
Sub SFaccRep()
'
' 4. Замена отчета SFacc по Организациям из первого листа файла MatchSF-1C
'   3.1.12
'  9.2.2012 - ревизия
' 30.4.12 - SFaccColFill для Платежей, Договоров и Организаций
    
    Dim LO, Ln As Integer           ' количества строк в старом и новом отчетах
    Dim Same As String
    
    ModStart SFacc, "Обновление листа отчета Salesforce по Организациям SFacc"

    LinesOld = Sheets(SFacc).UsedRange.Rows.count ' кол-во строк в старом отчете
    Lines = Sheets(1).UsedRange.Rows.count        ' кол-во строк в новом отчете
    LO = LinesOld - SFresLines
    Ln = Lines - SFresLines
    
    CheckSheet 1, Lines - 4, 1, SFaccRepName
    CheckSheet SFacc, LinesOld - 4, 1, SFaccRepName
    
    If LO = Ln Then Same = "тоже " Else Same = ""
    MsgBox "В новом отчете " & Ln - 1 & " Организаций, в прежнем " & Same & LO - 1
                                                            
    Sheets(SFacc).Rows("1:1").Copy Sheets(1).Rows("1:1")    ' копируем заголовок
    AutoFilterReset 1

    Sheets(SFacc).Name = "Tmp"          ' переименовываем прежний отчет в "Tmp"
    Sheets(1).Name = SFacc              '   .. а новый в SFacc

' заменяем ссылки в фоSheetNmрмулах других листов
    Call SheetsCtrlH(2, "Tmp!", "SFacc!")
    Call SFaccCol(2, 3)                ' лист Платежей
   
    Call SheetsCtrlH(4, "Tmp!", "SFacc!")
    Call SFaccCol(4, 7)                ' лист Договоров

    Call SheetsCtrlH("DogovorHeader", "Tmp!", "SFacc!") ' шаблон Dogovor_Insert

    Call SheetsCtrlH("Список клиентов 1C", "Tmp!", "SFacc!")
    Call SFaccCol("Список клиентов 1C", 3) ' отчет по Справочнику Клиентов 1С
   
    Sheets(1).Move Before:=Sheets("Tmp")    ' перемещаем новый отчет перед прежним
    Sheets("Tmp").Delete                    ' уничтожаем прежний отчет SF
    Sheets(SFacc).Tab.Color = rgbYellow     ' окрашиваем Tab нового отчета
    
    SFaccColFill PAY_SHEET  ' в колонке 1 если Организация есть в SF
    SFaccCol PAY_SHEET, PAY_RESLINES    ' раскрашиваем колонку A
    SFaccColFill DOG_SHEET  ' в колонке 1 если Организация есть в SF
    SFaccCol DOG_SHEET, DOGRES    ' раскрашиваем колонку A
    SFaccColFill Acc1C              ' в колонке 1 если Организация есть в SF
    SFaccCol Acc1C, ACC1C_RES       ' раскрашиваем колонку A

    ModEnd SFacc
End Sub
