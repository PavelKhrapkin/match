Attribute VB_Name = "Match1C"
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
    
    Public Const DownloadDir = "C:\Users\Пользователь\Downloads\"

' Названия листов - отчетов. Перечислено по порядку в Match.xlsm слева направо
'   !!! после заверщения параметризации можно этот список деклараций    !!!
'   !!! ... разнести по модулям                                         !!!
    Public PaidSheet As String  ' обычно отчет 1С по Договорам на первом месте,
                                    ' !!! но это надо параметризовать!!!
    Public DogSheet As String   ' обычно отчет 1С по Договорам на втором месте,
                                    ' !!! но это надо параметризовать!!!
    Public Const DogHeader = "DogovorHeader" ' шаблоны для DL Dogovor_Insert

'    Public Const PartnerCenter = "PartnerCenter"    ' имя листа отчета из
                                '                  PartnerCenter.Autodesk.com
    Public Const PaidContract = "P_PaidContract" ' рабочий лист- список новых
                                '                   .. оплаченных контрактов
    Public Const PaidNoContract = "P_PaidNoContract" ' список новых платежей
                                '               без контрактов - времянка!!!
'    Public Const PaidUpdate = "P_Update"    ' рабочий лист - список новых
                                '               платежей для DL - времянка!!!
    Public Lines As Integer     ' количество строк текущего/нового отчета
    Public LinesOld As Integer  ' количество строк старого отчета
    Public AllCol As Integer    ' Количество колонок в таблице отчета
    Public Doing As String      ' строка в Application.StatusBar - что делает модуль
Sub MoveToMatch()
Attribute MoveToMatch.VB_Description = "8.2.2012 - перемещение входного отчета на первый лист MatchSF-1C.xlsb,  распознавание его по штампу и запуск макроса по его замене "
Attribute MoveToMatch.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' 1. Macro to match SF and 1C Reports Pavel Khrapkin 23-Dec-2011
'
' Keyboard Shortcut: Ctrl+Shift+M
'
'       8.2.2012 - распознаем новый отчет, запускаем его обработку

    Dim SFrepName As String
    Const ErMsg = "На первом листе нераспознанный новый отчет"
    
    Sheets(1).Move Before:=Workbooks("Match SF-1C.xlsm").Sheets(1)
    
    Lines = EOL(1)
        
    If Cells(Lines, 1) = SFstamp Then                               '** отчет SF **
        Select Case Cells(Lines - 4, 1)
            Case SFpayRepName: Call Match1C_SF  ' отчет SF по Платежам
            Case SFcontrRepName: Call SFDreport ' отчет SFcontr - Договоры
            Case SFaccRepName: Call SFaccRep    ' отчет SFacc - Организации
            Case SFoppRepName: Call SFoppRep    ' отчет SFopp по Проектам
            Case SFadskRepName: ADSKfromSFrep   ' отчет SF по Autodesk
            Case SFpaRepName: SF_PA_Rep         ' отчет связкам Платежей с ADSK
            Case Else: MsgBox ErMsg & " Salesforce.", , "РАЗБЕРИСЬ!"
        End Select                                              '** отчеты 1С и Autodesk **
    ElseIf Cells(1, 1) = Stamp1Cpay1 And Cells(1, 2) = Stamp1Cpay2 Then From1Cpayment
    ElseIf Cells(1, 2) = Stamp1Cdog1 And Cells(1, 4) = Stamp1Cdog2 Then From1Cdogovor
    ElseIf Cells(1, 5) = Stamp1Cacc1 And Cells(1, 6) = Stamp1Cacc2 Then From1Caccount
    ElseIf Cells(1, 40) = StampADSKp1 And Cells(1, 42) = StampADSKp2 Then FrPartnerCenter
    Else: MsgBox ErMsg, , "РАЗБЕРИСЬ!"
    End If
    ActiveWorkbook.Save
End Sub
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
