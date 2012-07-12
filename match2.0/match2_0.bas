Attribute VB_Name = "match2_0"
'---------------------------------------------------------------------------
' Макросы для работы с файлом отчетов из 1С и Salesforce Match SF-1C.xlms
'
' * MoveToMatch    - перенос Листа на первое место Match1SF    (Ctrl/Shift/M)
' * TriggerOptionsFormulaStyle  - переключение моды A1/R1C1    (Ctrl/Shift/R)
'
' П.Л.Храпкин 2.1.2012
'   28.1.2012 - работы по параметризации имен и позиций листов
'    5.2.2012 - в MoveToMatch - распознавание входного отчета по штампу
'   16.5.2012 - добавлен отчет SF_PA
'    2.6.2012 - TriggerOptionsFormulaStyle A1/R1C1

    Option Explicit    ' Force explicit variable declaration
    
''''    Public Const DownloadDir = "C:\Users\Пользователь\Downloads\"
''''
''''' Названия листов - отчетов. Перечислено по порядку в Match.xlsm слева направо
'''''   !!! после заверщения параметризации можно этот список деклараций    !!!
'''''   !!! ... разнести по модулям                                         !!!
''''    Public PaidSheet As String  ' обычно отчет 1С по Договорам на первом месте,
''''                                    ' !!! но это надо параметризовать!!!
''''    Public DogSheet As String   ' обычно отчет 1С по Договорам на втором месте,
''''                                    ' !!! но это надо параметризовать!!!
''''    Public Const DogHeader = "DogovorHeader" ' шаблоны для DL Dogovor_Insert
''''
'''''    Public Const PartnerCenter = "PartnerCenter"    ' имя листа отчета из
''''                                '                  PartnerCenter.Autodesk.com
''''    Public Const PaidContract = "P_PaidContract" ' рабочий лист- список новых
''''                                '                   .. оплаченных контрактов
''''    Public Const PaidNoContract = "P_PaidNoContract" ' список новых платежей
''''                                '               без контрактов - времянка!!!
'''''    Public Const PaidUpdate = "P_Update"    ' рабочий лист - список новых
''''                                '               платежей для DL - времянка!!!
Sub MoveToMatch()
Attribute MoveToMatch.VB_Description = "8.2.2012 - перемещение входного отчета на первый лист MatchSF-1C.xlsb,  распознавание его по штампу и запуск макроса по его замене "
Attribute MoveToMatch.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' <*> MoveToMatch() - перемещение входного отчета в базу и запуск его обработки
'
' Keyboard Shortcut: Ctrl+Shift+M
'
'Pavel Khrapkin 23-Dec-2011
' 8.2.2012 - распознаем новый отчет, запускаем его обработку
' 11.7.12 - match2.0 - распознавание отчета, перенос его в один из файлов базы и запуск обработки

    Dim NewRep As String            ' имя файла с новым отчетом
    Dim RepName As String           ' имя нового отчета
    
    NewRep = ActiveWorkbook.Name
    Lines = EOL(1, Workbooks(NewRep))

    Set DB_MATCH = Workbooks.Open(F_MATCH, UpdateLinks:=False)
    
    Dim iDBs As Integer         'параметр TOCmatch - количество баз данных
    iDBs = DB_MATCH.Sheets(1).Cells(4, TOC_PAR_1_COL)
    
'------ распознавание Штампа NewRep по таблице TOCmatch -------------
    Dim TOCline As Range        '= строка TOC match
    With TOCline
        For Each TOCline In Range(Cells(5, 1), Cells(5 + iDBs, BIG)).Rows
            If IsRightStamp(TOCline, NewRep) Then GoTo RepNameCheck
        Next TOCline
        GoTo FatalNewRep
RepNameCheck:
        Dim FrTOC As Integer, ToTOC As Integer  'строки поиска RepName в TOC
        FrTOC = .Cells(1, TOC_PAR_2_COL)
        ToTOC = .Cells(1, TOC_PAR_3_COL)
        For Each TOCline In Range(Cells(FrTOC, 1), Cells(ToTOC, TOC_PAR_3_COL)).Rows
            If IsRightStamp(TOCline, NewRep) Then GoTo RepNameHandle
        Next TOCline
        GoTo FatalNewRep
RepNameHandle:
        
    End With
        
    If InStr(LCase$(Cells(Lines + 3, 1)), "salesforce.com") <> 0 _
            And Cells(Lines + SFresLines, 1) = SFstamp Then
            
        Set DB_SFDC = Workbooks.Open(F_SFDC, UpdateLinks:=False)
        Workbooks(NewRep).Sheets(1).Move Before:=DB_SFDC.Sheets(1)
        RepName = Cells(Lines + 2, 1)
        Lines = Lines - SFresLines
        
        Select Case RepName
            Case SFpayRepName:
                Application.Run ("SFDC.xlsm!Match1C_SF")    ' отчет SF по Платежам
            Case SFcontrRepName:
                Application.Run ("SFDC.xlsm!SFDreport")     ' отчет SFcontr - Договоры
            Case SFaccRepName:
                Application.Run ("SFDC.xlsm!SFaccRep")      ' отчет SFacc - Организации
            Case SFcontactRepName:
                Application.Run ("SFDC.xlsm!SFcontactRep")  ' отчет SFcont по Контактам
            Case SFoppRepName:
                Application.Run ("SFDC.xlsm!SFoppRep")      ' отчет SFopp по Проектам
            Case SFadskRepName:
                Application.Run ("SFDC.xlsm!ADSKfromSFrep") ' отчет SF по Autodesk
            Case SFpaRepName:
                Application.Run ("SFDC.xlsm!SF_PA_Rep")     ' отчет по связкам Платежей с ADSK
            Case Else:
                ErrMsg FATAL_ERR, "Не распознан отчет Salesforce.com"
        End Select

        
        '** отчеты 1С и Autodesk **
    ElseIf Cells(1, 1) = Stamp1Cpay1 And Cells(1, 2) = Stamp1Cpay2 Then
        Application.Run ("1C.xlsm!From1Cpayment")    ' отчет 1С по Платежам
    ElseIf Cells(1, 2) = Stamp1Cdog1 And Cells(1, 4) = Stamp1Cdog2 Then
        Application.Run ("1C.xlsm!From1Cdogovor")    ' отчет 1С по Договорам
    ElseIf Cells(1, 5) = Stamp1Cacc1 And Cells(1, 6) = Stamp1Cacc2 Then
        Application.Run ("1C.xlsm!From1Caccount")    ' отчет 1С по Клиентам
'''    ElseIf Cells(1, 40) = StampADSKp1 And Cells(1, 42) = StampADSKp2 Then
'''        FrPartnerCenter
    Else: GoTo FatalNewRep
        End
    End If
        
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    Workbooks(NewRep).Close

'------------- match TOC write -------------------------
    
'''    With DB_MATCH
'''            If TOCline.Cells(1, TOC_REPNAME_COL) = RepName Then
'''                With TOCline
'''                    .Cells(1, TOC_LOAD_COL) = Now
'''                    .Cells(1, TOC_HANDLE_COL) = ""
'''                    .Cells(1, TOC_EOL_COL) = Lines
'''                End With
'''                Exit For
'''            End If
'''        .Cells(1, 1) = Now
'''        .Save
'''    End With
    Exit Sub
FatalNewRep:
    ErrMsg FATAL_ERR, "Входной отчет '" & NewRep & "' не распознан"
End Sub
Function IsRightStamp(TOCline, NewRep) As Boolean
'
' - IsRightStamp(TOCline) - проверка правильности штампа в NewRep по строке TOCline
' 12.7.2012

    Dim NewRepStamp As String       ' штамп нового отчета
    
    Dim Stamp As String         '= строка - штамп
    Dim StampType As String     'тип штампа: строка (=) или подстрока
    Dim Stamp_R As Integer      'номер строки, где штамп
    Dim Stamp_C As Integer      'номер колонки, где штамп
    Dim ParCheck As Integer     'параметр TOCmatch - строка дополнительной проверки штампа
    
    IsRightStamp = False
        
    With TOCline
        Do
            Stamp = .Cells(1, TOC_STAMP_COL)
            If Stamp = "" Then Exit Function        ' отсутствует штамп - не годится!
            StampType = .Cells(1, TOC_STAMP_TYPE_COL)
            Stamp_R = .Cells(1, TOC_STAMP_R_COL)
            If .Cells(1, TOC_EOL_COL) = "EOL" Then Stamp_R = Stamp_R + Lines
            Stamp_C = .Cells(1, TOC_STAMP_C_COL)
            NewRepStamp = Workbooks(NewRep).Sheets(1).Cells(Stamp_R, Stamp_C)
            
            If StampType = "=" And NewRepStamp <> Stamp Then
                Exit Function
            ElseIf StampType = "I" And InStr(LCase$(NewRepStamp), LCase$(Stamp)) = 0 Then
                Exit Function
            Else: If StampType <> "=" And StampType <> "I" Then _
                ErrMsg FATAL_ERR, "Сбой в структоре TOCmatch: тип штампа =" & StampType
            End If
        
            ParCheck = .Cells(1, TOC_PAR_1_COL)
            If IsNumeric(ParCheck) And ParCheck > 0 Then
                Set TOCline = Range(Cells(ParCheck, 1), Cells(ParCheck, BIG))
            End If
        Loop While ParCheck <> 0
    End With

    IsRightStamp = True

End Function
Sub TriggerOptionsFormulaStyle()
'
' * Trigger Options-Formula Style A1/R1C1
'
' CTRL+Shift+R
'
' 2.6.12
    If Application.ReferenceStyle = xlR1C1 Then
        Application.ReferenceStyle = xlA1
    Else
        Application.ReferenceStyle = xlR1C1
    End If
End Sub
