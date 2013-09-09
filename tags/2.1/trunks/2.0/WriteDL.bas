Attribute VB_Name = "WriteDL"
'================================================================================
' Модуль для загрузки данных посредством Informatica Data Loader
'
' * NewContractDL  - в отчете 1С появился новый Договор - запись в Dogovor.csv
' - PaymentUpd(PayKod, ContrId) - Update Платежа - запись в SF через Data Loader
' - NewContract(Dogovor, MainDog, ContrK) - создание нового договора ContrK в SF
' - DogFormat(Wsheet) - форматирование листа Wsheet для вывода в Dogovor.csv
' S WrCSV(SheetN, Directory, FileToWrite, Bat) - Шаг - запись CSV
' - WriteCSV(SheetN, FileName, ..)  - запись листа SheetN в файл для загрузки DL
'
'   23.9.2012 - выделение модуля AddressParse
'   31.10.12

Option Explicit

Const ACC_NEWDOG_COL = 7    ' колонка Клиент

Const C1C = 3       ' Номер листа отчета по Договорам из 1С
Dim C1Cname         ' Имя листа Договоров, оно включает дату

Sub NewContractDL()
'
' подготовка файла Dogovor.csv по для DL Task "Dogovor_Insert"
'   24.1.2012

    Const ResLines = 7      ' число строк сводки под отчетом по Договорам (пятка)

    Dim MyRange As Range
    Dim W, L, Wsheet
    
    EOL_SFD = EOL(SFD) - SFresLines
    
    Sheets(3).Select
    W = ActiveSheet.UsedRange.Columns.Count
    L = ActiveSheet.UsedRange.Rows.Count - ResLines
    
    Sheets(3).Select
    Set MyRange = Range(Cells(1, 1), Cells(L, W))
    MyRange.AutoFilter Field:=1, Criteria1:="<>"            ' Организация есть в SF
    MyRange.AutoFilter Field:=2, Criteria1:="=Нет в SF"     ' Договора нет в SF
    Range(Cells(1, 10), Cells(L, 23)).Copy
    
    Sheets.Add After:=Sheets(Sheets.Count)  ' создаем новый лист - временный
    Wsheet = ActiveSheet.Name
    ActiveSheet.Paste
    
    Call DogFormat(Wsheet)  ' форматируем рабочий лист Wsheet для вывода Договора
    
    ChDir "C:\Users\salesforce\Dogovor"
    WriteCSV Wsheet, "Dogovor.txt"

    Shell "quota2.bat"
    ActiveSheet.Delete                      ' уничтожаем времменый лист
End Sub
Sub PaymentUpd(PayKod, ContrId)
'
' Update Платежа - запись в SF через Data Loader
' На сегодняшний день обновляем только поле - ссылку на Договор
'   11.3.12

    Dim PaymentId As String

    PaymentId = ""
    On Error Resume Next
    PaymentId = WorksheetFunction.VLookup(PayKod, _
        Sheets(SF).Range("B:R"), 17, False)
    On Error GoTo 0
    If PaymentId = "" Then GoTo Err
    EOL_PaymentUpd = EOL_PaymentUpd + 1
    Sheets(P_PaymentUpd).Cells(EOL_PaymentUpd, 1) = PaymentId
    Sheets(P_PaymentUpd).Cells(EOL_PaymentUpd, 2) = ContrId
    Exit Sub
Err:
    MsgBox "Не найден Платеж SF по коду '" & PayKod & "'", , "FATAL ERROR!"
    Stop
End Sub
Sub NewContractXX(Dogovor, MainDog, ContrK)
'
' XXX создание нового договора ContrK в SF
'   31.3.12

    Dim StartDate, OurLegalName, IdWe, Salesman, IdOwner As String
    Dim Client, DogType, Value, DogCurrency, CurCod, Dat As String

    Dim i
       
' выборка данных Договора из отчета 1С
    With Sheets(DOG_SHEET)
        For i = 2 To EOL_DogSheet
            If ContrK = .Cells(i, DOGCOD_COL) Then
                StartDate = .Cells(i, DOG1CDAT_COL)
                OurLegalName = .Cells(i, DOG1CFIRM_COL)
                IdWe = WorksheetFunction.VLookup(OurLegalName, Range("Мы"), 2, False)
                Salesman = .Cells(i, DOG1CSALE_COL)
                IdOwner = ""
                On Error Resume Next
                IdOwner = WorksheetFunction.VLookup(Salesman, Range("Продавец_в_SF"), 3, False)
                On Error GoTo 0
                If IdOwner = "" Then IdOwner = "00520000001460C"    ' Все ВСЕ
                Client = .Cells(i, DOG1CACC_COL)
                DogType = .Cells(i, DOGTYPE_COL)
                Value = Dec(.Cells(i, DOGSUM1C_COL))
                DogCurrency = .Cells(i, DOGCUR1C_COL)
                Dat = DDMMYYYY(.Cells(i, DOG1CDAT_COL))
                Exit For
            End If
        Next i
    End With
    
' Поиск Id Основного договора в SFD
    Dim IdMainDog As String
    Dim X() As String
    
    IdMainDog = ""
    If MainDog <> "" Then
        X = split(Compressor(MainDog), " ")
        On Error Resume Next
        IdMainDog = WorksheetFunction.VLookup(X(1), _
            Sheets(SFD).Range("B:O"), 14, False)
        On Error GoTo 0
        If IdMainDog = "" Then
            LogWr "NewContract: не найден Основной договор по '" _
                & ContrK & "' -- Договор в SF не создаем!"
            Exit Sub
        End If
    End If

' заполнение полей адреса для счета по SFacc
    Dim Street, City, State, Index, Country As String
    With DB_SFDC.Sheets(SFacc)
        For i = 2 To EOL_SFacc
            If .Cells(i, SFACC_ACC1C_COL) = Client Then
                Street = .Cells(i, SFACC_STREET_COL)
                City = .Cells(i, SFACC_CITY_COL)
                State = .Cells(i, SFACC_STATE_COL)
                Index = .Cells(i, SFACC_INDEX_COL)
                Country = .Cells(i, SFACC_COUNTRY_COL)
                Exit For
            End If
        Next i
    End With
    
' Заполнение листа NewContract данными по новому договорут для CSV
    EOL_NewContr = EOL_NewContr + 1
    With DB_MATCH.Sheets(NewContract)
        .Cells(EOL_NewContr, NEWDOG_DOGOVOR_COL) = " " & Dogovor
        .Cells(EOL_NewContr, NEWDOG_DATE_COL) = DDMMYYYY(StartDate)
        .Cells(EOL_NewContr, NEWDOG_WE_COL) = OurLegalName
        .Cells(EOL_NewContr, NEWDOG_IDWE_COL) = IdWe
        .Cells(EOL_NewContr, NEWDOG_SALESMAN_COL) = Salesman
        .Cells(EOL_NewContr, NEWDOG_IDOWNER_COL) = IdOwner
        .Cells(EOL_NewContr, NEWDOG_CLIENT_COL) = Client
        .Cells(EOL_NewContr, NEWDOG_IDACC_COL) = AccId(Client)
        .Cells(EOL_NewContr, NEWDOG_TYPE_COL) = DogType
        .Cells(EOL_NewContr, NEWDOG_MAINDOG_COL) = IdMainDog
        .Cells(EOL_NewContr, NEWDOG_VALUE_COL) = Value
        .Cells(EOL_NewContr, NEWDOG_CURRENCY_COL) = DogCurrency
        CurCod = CurISO(DogCurrency)
        .Cells(EOL_NewContr, NEWDOG_CURRENCYCOD_COL) = CurCod
        .Cells(EOL_NewContr, NEWDOG_RATE_COL) = Dec(CurRate(CurCod))
        .Cells(EOL_NewContr, NEWDOG_RATEDATE_COL) = Dat
        .Cells(EOL_NewContr, NEWDOG_ACCSTREET_COL) = Street
        .Cells(EOL_NewContr, NEWDOG_ACCCITY_COL) = City
        .Cells(EOL_NewContr, NEWDOG_ACCSTATE_COL) = State
        .Cells(EOL_NewContr, NEWDOG_ACCINDEX_COL) = Index
        .Cells(EOL_NewContr, NEWDOG_ACCCOUNTRY_COL) = Country
    End With
    Exit Sub
Err:    MsgBox "NewContract '" & ContrK & "'", , "ERROR!"
    Stop
End Sub
Sub ContrLink(ContrCode, OppId)
'
' связывание Договора (код ContrK) с Проектом (OppId)
'   18.3.12


    EOL_ContrLnk = EOL_ContrLnk + 1
    With Sheets(NewContractLnk)
        .Cells(EOL_ContrLnk, 1) = ContractId(ContrCode)
        .Cells(EOL_ContrLnk, 2) = OppId
    End With
End Sub
Sub DogFormat(Wsheet)
'
' подпрограмма форматирования рабочего листа Wsheet для вывода в Dogovor.csv
'   25.1.2012
'   7.3.12 - убираем лишние пробелы в имени клиента
'   14.3.12 - работа с Основным Договором
    
    Const DogHdr = "DogovorHeader"  ' шаблон Договора для DL Dogovor_Insert
    Dim MyRange As Range
    Dim L, i
    
    Sheets(DogHdr).Rows("1:3").Copy
    Sheets(Wsheet).Select
    Lines = ActiveSheet.UsedRange.Rows.Count
    L = Lines + 3
    Set MyRange = Range("A1:U" & L)
    
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown     ' строки из шаблона
            
    Columns("B:B").NumberFormat = "@"   ' заменяем Дату на текст
    For i = 5 To L
        Cells(i, 2) = Day(Cells(i, 3)) & "-" & Month(Cells(i, 3)) & "-" & Year(Cells(i, 3))
    Next i
    Range("C4:C" & L).Delete Shift:=xlToLeft     ' B - Delete - Дата подписания Договора
    
 ' вставляем колонку D - IdWe и заполняем ее из шаблона VLOOKUP
    Range("D4:D" & L).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D3").AutoFill Destination:=Range("D3:D" & L), Type:=xlFillDefault
    
    For i = 5 To L
        Cells(i, 6) = Compressor(Cells(i, 6)) ' убираем лишние пробелы в имени клиента
    Next i

 ' вставляем колонку F - IdOwner и заполняем F и H формулами из шаблона
    Range("F4:F" & L).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F3").AutoFill Destination:=Range("F3:F" & L), Type:=xlFillDefault
    Range("H3").AutoFill Destination:=Range("H3:H" & L), Type:=xlFillDefault

    Columns("L:L").NumberFormat = "@"   ' заменяем десятичную запятую на точку в Сумме
    Call MainContrIdSet(Wsheet, L)      ' заменяем поле Основной контракт на Id из SFD
    
    For i = 5 To L
        Cells(i, 12) = Replace(Cells(i, 12), ",", ".")
    Next i

    Range("M4:N" & L).Delete Shift:=xlToLeft    ' убираем колонки Сумм по счетам и платежам

    Range("N3:U" & L).FillDown          ' вставляем формулы из шаблона по остальным полям
    Rows("2:4").Delete                  ' стриаем строка из шаблона
    Rows("1:1").AutoFilter
    MyRange.AutoFilter Field:=4, Criteria1:="<>#N/A"
    MyRange.AutoFilter Field:=6, Criteria1:="<>#N/A"
    MyRange.AutoFilter Field:=8, Criteria1:="<>#N/A"
    MyRange.AutoFilter Field:=14, Criteria1:="<>#N/A"
    
    Range("A1:U" & Lines).Copy Cells(Lines + 2, 1)  ' копируем Договоры, в которых нет #N/A
    Call AutoFilterReset(Wsheet)
    Rows("1:" & Lines + 1).Delete

    MsgBox "Во избежание дублирования Договоров" & vbCrLf & _
            "     не забудь про Отчет SF - SFD!", vbExclamation, "ВНИМАНИЕ!"
End Sub
Sub MainContrIdSet(Wsheet, L)
'
' определение Id Основного контракта как он указан в отчете по Договорам 1С по SFD
'   14.3.12

    Const MAIN_DOG_SFD_COL = 5      ' колонка "Осн.Договор"
    Const ACC_SFD_COL = 6           ' колонка "Имя организации 1С" в SFD
    Const MAIN_DOGID_SFD_COL = 19   ' колонка "Код основного договора"
    
    Dim i, j, IsMain, X
    Dim MainContr, MainContrId, Account As String
    
    For i = 5 To L
        MainContr = Sheets(Wsheet).Cells(i, 10)
        If MainContr <> "" Then
            MainContrId = ""
            Account = Sheets(Wsheet).Cells(i, ACC_NEWDOG_COL)
            With Sheets(SFD)
                For j = 2 To EOL_SFD
                    X = .Cells(j, MAIN_DOG_SFD_COL)
                    IsMain = InStr(MainContr, X)
                    If IsMain <> 0 And X <> "" Then
                        If .Cells(j, ACC_SFD_COL) = Account Then
                            MainContrId = .Cells(j, MAIN_DOGID_SFD_COL)
                            GoTo FoundMainContr
                        End If
                    End If
                Next j
            End With
            LogWr "     WARNING:  не найден в SFD Основной Договор '" _
                & MainContr & "'"
            MainContrId = "НЕ НАЙДЕН ОСНОВНОЙ ДОГОВОР!"
FoundMainContr:
            Sheets(Wsheet).Cells(i, 10) = MainContrId
        End If
    Next i
End Sub
Sub WrCSV(SheetN As String, Directory, FileToWrite, Bat)
'
' S WrCSV(SheetN, Directory, FileToWrite, Bat) - Шаг - запись CSV
'   23.9.12
'   30.9.12 - запись числа новых строк в SheetN в WrProcResult
'   19.10.12 - использование WP_TMP
'   31.10/12 - не стираем лист CSV

    Dim NewLines As Long
    
    Set DB_TMP = FileOpen(F_TMP)
    DB_TMP.Sheets(SheetN).Activate
    NewLines = EOL(SheetN, DB_TMP)
    If NewLines <= 1 Then
'        Application.DisplayAlerts = False
'        ActiveSheet.Delete
'        Application.DisplayAlerts = True
    Else
        ChDir Directory
        WriteCSV SheetN, FileToWrite
        Shell Bat
        LogWr "WrCSV> в '" & FileToWrite & "' записано " & NewLines & " строк."
    End If
    
    WrProcResult NewLines - 1
End Sub
Sub WriteCSV(SheetN, FileName, _
    Optional Row0 As Integer = 1, Optional Col0 As Integer = 1, _
    Optional RowLast As Integer, Optional ColLast As Long)
'
' запись листа SheetN в файл FileName в формате CSV Tab Delimited и его стирание
'   24.1.2012
'    3.2.2012 Optionan Range записываемого массива
'   13.2.2012 EOL(SheetN), чтобы не писать пустые строки в конце
'   29.2.2012 bug fix - Cells без Sheets приводил в выводу неправильных листов
'   23.9.12 - отладка в match 2.0
    
    Dim S, Col
    
    Open FileName For Output As #1
    
    If RowLast = 0 Then RowLast = EOL(SheetN)
    If ColLast = 0 Then
        ColLast = ActiveSheet.UsedRange.Columns.Count
    End If
    For S = Row0 To RowLast
        For Col = Col0 To ColLast
            Print #1, ActiveSheet.Cells(S, Col);
            If Col <> ColLast Then Print #1, vbTab;
        Next Col
        Print #1, " "
    Next S
    Close #1
'    MsgBox "Записано " & (RowLast - Row0 + 1) & " строк" & vbCrLf & "в файл " & Chr(171) & FileName & Chr(187)
End Sub

Sub WritePaid(FileName)
'
' подпрограмма записи файл FileName для процедуры DL Payment
'   10.1.2012
    
    Range("C1:I" & ActiveSheet.UsedRange.Rows.Count).Copy   ' копируем нужные колонки из P_Update
    Sheets.Add After:=Sheets(Sheets.Count)  ' создаем новый лист - временный
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False           ' помещаем в него из P_Update Paste Value

    Open FileName For Output As #1
    
    With ActiveSheet.UsedRange
        For i = 1 To .Rows.Count
            Print #1, Range("A" & i) & vbTab & Range("B" & i) & vbTab & Range("C" & i) & vbTab & _
                Range("D" & i) & vbTab & Range("E" & i) & vbTab & Range("F" & i) & vbTab & Range("G" & i)
        Next
    End With
    Close #1
    ActiveSheet.Delete                      ' уничтожаем временный лист
End Sub



