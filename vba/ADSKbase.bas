Attribute VB_Name = "ADSKbase"
'---------------------------------------------------------------------------------
' ADSKbase  - работа с базой данных в ADSK.xlsx
'
'[*] SubscriptionsADSKpass      - проход по отчету ADSK Subscriptions
' - IsContrADSKinSF(ContrADSK)  - возвращает TRUE если контракт ContrADSK есть в SF
'(*) ADSK_GFP_Upgrade()         - проход по отчету ADSK GFP Upgrades
' - IsSN_OK(iADSK, iSF)         - TRUE если атрибуты SN_SF соответствуют SN_ADSK
' - IsSNitemOK(iMap, SN_SFitem, SN_ADSKitem) - проверки соотвествия отдельных арибутов SN
' - ErrSN(iADSK, ColADSK, iSF, ColSF) - сообщение о несоответствии и раскраска ADSKrep
' - ContrADSKinSFatr(ContrADSK, iSF) - возвращает атрибуты котракта по ADSKinSF
' - AccNinSFatr(AccN, iSF) - находит AccN в ADSKinSF и помещает в SN_SF
' - ZeroSNatr() - возвращаент инициализированную структуру SNatr
' - SNinSFatr(SN, iSF)          - возвращает атрибуты SN в SF по ADSKinSF
' - SNinADSKatr(iADSK)          - возвращает атрибуты SN по строке iADSK отчета ADSKrep
' - SNvalByMap(iADSK, ColADSK, [ValType]) - возвращает значете поля типа ValType
'
'   8.6.2012

Option Explicit
Sub ADSK_TOC_FormOutput()
'
' (*) ADSK_TOC_FormOutput() - вывод формы по Содержанию (TOC) ADSK.xlsx
'   8.6.12

    Dim S As Range
    
'---- вначале загружем Оглавление из ADSK.xlsx
    Workbooks.Open ("C:\Users\Пользователь\Desktop\Работа с Match\SFconstrTMP\ADSK\ADSK.xlsx")
    Windows("ADSK.xlsx").Activate
'    Sheets(TOC_ADSK).Select
'    Sheets(TOC_ADSK).Copy Before:=Workbooks("Match SF-1C.xlsm").Sheets(We)

    ADSK_TOC_Form.TOClist.RowSource = ""
    ADSK_TOC_Form.TOClist.ColumnCount = 2
    
    For Each S In Sheets(TOC_ADSK).Range("TOC_ADSK_Range").Rows
'        E = S.Cells(1, WE_ERR_COL)  ' число несоответствий для данного Продавца
'        If E > 0 Then
            ADSK_TOC_Form.TOClist.AddItem S.Cells(1, 3).value
'            ADSK_TOC_Form.TOClist.List(N - 1, 1) = E
'            N = N + 1
        End If
    Next S
    CheckingForm.Show
End Sub
Sub SubscriptionsADSKpass()
'
' [*] SubscriptionsADSKpass - проход по отчету ADSK Subscriptions
'
' загружаем отчет Subscriptions из базы ADSK.xlsx
' 1) проверяем, есть ли Контракт ADSK в SF. Если нет - NewContrADSK
' 2) проверяем, есть ли SN в SF. Если нет - NewSN(ContrADSK)
' 3) проверяем, есть ли Проект - продление подписки
' если все это есть - стираем лист Subscriptions
'
'   27.5.12

    Const RepName = "Subscriptions"
    
    Const SBS_CONTR_COL = 13    'поле "Agreement Number" - номер Контракта ADSK
    Const SBS_SN_COL = 16       'поле "Subs Serial #" - SN ADSK
    
    Dim EOLsbs As Integer
    Dim ContrADSK As String     '= значение поля SBS_CONTR - Контракт ADSK
    Dim SN As String            '= значение поля SN - Серийный номер подписки
    Dim NoContr As Integer      '= число найденных Контрактов, которых нет в SF
    Dim NoSN As Integer         '= число найденных SN, которых нет в SF
    Dim i As Integer
    
    ModStart ADSKfrSF, "Проход по отчету Autodesk 'Subscriptions'"
    CheckSheet ADSKfrSF, EOL_ADSKfrSF + 2, 1, ADSKfrSFstamp
    
    GetSheetFrADSK RepName

    With Sheets(RepName)
        For i = 3 To EOL_ADSK
    '--- есть Контракт ADSK в SF?
            ContrADSK = .Cells(i, SBS_CONTR_COL)
            If IsContrADSKinSF(ContrADSK) Then
                GoTo SNcheck
            Else
                .Cells(i, SBS_CONTR_COL).Interior.Color = rgbRed
                NoContr = NoContr + 1
                GoTo NXT
            End If
            
    '--- есть SN ADSK в SF?
SNcheck:    SN = .Cells(i, SBS_SN_COL)
'            If IsSNinSF(SN, ContrADSK) Then
''                   If IsOpp(SN, ContrADSK) Then
'            Else
'                .Cells(i, SBS_SN_COL).Interior.Color = rgbRed
'                NoSN = NoSN + 1
'            End If
NXT:    Next i
    End With
    
    MsgBox "В SF отсутствует " & NoContr & "контактов ADSK"
    ModEnd ADSKfrSF
End Sub
Function IsContrADSKinSF(ContrADSK) As Boolean
'
' - IsContrADSKinSF(ContrADSK) - возвращает TRUE если контракт ContrADSK есть в SF
'   28.5.12

    Dim ContrN As String
    
    IsContrADSKinSF = False
    
    ContrN = ""
    On Error Resume Next
    ContrN = WorksheetFunction.VLookup(ContrADSK, _
        Sheets(ADSKfrSF).Range("A:A"), 1, False)
    On Error GoTo 0
    
    If ContrN <> "" Then IsContrADSKinSF = True
    
End Function
Sub ADSK_GFP_Upgrade()
'
' (*) ADSK_GFP_Upgrade() - проход по отчету ADSK GFP
'
' загружаем отчет GFP из базы ADSK.xlsx
' 1) проверяем, есть ли Контракт ADSK в SF. Если нет - NewContrADSK
' 2) проверяем, есть ли SN в SF. Если нет - NewSN(ContrADSK)
' 3) проверяем, есть ли Проект - продление подписки
' если все это есть - стираем лист Subscriptions
'
'   3.6.12
       
    Dim BadSN As Integer        '= счетчик неверных записей в SF по SN
    Dim iSF As Integer          '= номер строки ADSKfrSF, где найден SN
    Dim i As Integer
    
    ModStart ADSKfrSF, "Проход по отчету Autodesk GFP - Upgrade"
    CheckSheet ADSKfrSF, EOL_ADSKfrSF + 2, 1, ADSKfrSFstamp
    
'    GetSheetFrADSK "GFP"      '*** извлекаем отчет GFP из ADSK.xlsx
    GetSheetFrADSK "Subscription Extracts Coverage"      '*** извлекаем отчет из ADSK.xlsx

    BadSN = 0
       
    For i = 2 To EOL_ADSK
        SN_ADSK = SNinADSKatr(i)
        SN_SF = SNinSFatr(SN_ADSK.SN, iSF)
        If Not IsSN_OK(i, iSF) Then BadSN = BadSN + 1
    Next i
    
    ErrMsg TYPE_ERR, "В SF найдено " & BadSN & " строк по SN"
    ModEnd ADSKfrSF
End Sub
Function IsSN_OK(ByVal iADSK As Integer, ByVal iSF As Integer) As Boolean
'
' - IsSN_OK(iADSK, iSF, ErrCol)   - TRUE если атрибуты SN_SF соответствуют SN_ADSK
'               Если не соответствует - окрашивает ячейку iADSK в RepName и возвращает
'               номер первой ошибочной колонки
'
'       Атрибуты SN: номер контракта ADSK, AccN, AccName, Seats, Status и др
'       определены в Declarations.
'       Возвращает TRUE если атрибуты в SF соответствуют входным значениям SNatr
'       или не заданы (="").
'
'   5.6.12
    
    IsSN_OK = False
    
    If SN_SF.SN <> SN_ADSK.SN Or SN_SF.ErrFlag Then  ' SN есть в SF?
        ErrSN iADSK, ADSK_SN_COL
        ContrADSKinSFatr SN_ADSK.Contr, iSF         ' Конракт ADSK есть?
    End If
    
    If Not IsSNitemOK(ADSK_CONTR_COL, SN_SF.Contr, SN_ADSK.Contr) Then
        ErrSN iADSK, ADSK_CONTR_COL, iSF, SFADSK_CONTRACT_COL
        AccNinSFatr SN_ADSK.AccN, iSF               ' есть хотя бы AccN?
    End If
    If Not IsSNitemOK(ADSK_ACCN_COL, SN_SF.AccN, SN_ADSK.AccN) Then _
        ErrSN iADSK, ADSK_ACCN_COL, iSF, SFADSK_ACCN_COL
    If Not IsSNitemOK(ADSK_ACCNAME_COL, SN_SF.AccName, SN_ADSK.AccName) Then _
        ErrSN iADSK, ADSK_ACCNAME_COL, iSF, SFADSK_ACCNAME_COL
    If Not IsSNitemOK(ADSK_C_START_COL, SN_SF.C_Start, SN_ADSK.C_Start) Then _
        ErrSN iADSK, ADSK_C_START_COL, iSF, SFADSK_CONTRSTARTDAT_COL
    If Not IsSNitemOK(ADSK_C_END_COL, SN_SF.C_End, SN_ADSK.C_End) Then _
        ErrSN iADSK, ADSK_C_END_COL, iSF, SFADSK_CONTRENDDAT_COL
    If Not IsSNitemOK(ADSK_C_STAT_COL, SN_SF.C_Status, SN_ADSK.C_Status) Then _
        ErrSN iADSK, ADSK_C_STAT_COL, iSF, SFADSK_CONTRSTAT_COL
    If Not IsSNitemOK(ADSK_CM_NAME_COL, SN_SF.CM_Name, SN_ADSK.CM_Name) Then _
        ErrSN iADSK, ADSK_CM_NAME_COL, iSF, SFADSK_CM_NAME_COL
    If Not IsSNitemOK(ADSK_CM_MAIL_COL, SN_SF.CM_Mail, SN_ADSK.CM_Mail) Then _
        ErrSN iADSK, ADSK_CM_MAIL_COL, iSF, SFADSK_CM_MAIL_COL
    If Not IsSNitemOK(ADSK_CM_TEL_COL, SN_SF.CM_Tel, SN_ADSK.CM_Tel) Then _
        ErrSN iADSK, ADSK_CM_TEL_COL, iSF, SFADSK_CM_MAIL_COL
    If Not IsSNitemOK(ADSK_SN_DESCR_COL, SN_SF.Description, SN_ADSK.Description) Then _
        ErrSN iADSK, ADSK_SN_DESCR_COL, iSF, SFADSK_DESCRIPTION_COL
    If Not IsSNitemOK(ADSK_SEATS_COL, SN_SF.Seats, SN_ADSK.Seats) Then _
        ErrSN iADSK, ADSK_SEATS_COL, iSF, SFADSK_SEATS_COL
    If Not IsSNitemOK(ADSK_SN_STATUS_COL, SN_SF.Status, SN_ADSK.Status) Then _
        ErrSN iADSK, ADSK_SN_STATUS_COL, iSF, SFADSK_STATUS_SN_COL
    If Not IsSNitemOK(ADSK_DEPL_COL, SN_SF.Deployment, SN_ADSK.Deployment) Then _
        ErrSN iADSK, ADSK_DEPL_COL, iSF, SFADSK_DEPLOYMENT_COL
    If Not IsSNitemOK(ADSK_SBS_COL, SN_SF.Sbs, SN_ADSK.Sbs) Then _
        ErrSN iADSK, ADSK_SBS_COL, iSF, SFADSK_ISSBS_COL
'    If Not IsSNitemOK(ADSK_YEAR_COL, SN_SF.Release_Year, SN_ADSK.Release_Year) Then _
'        ErrSN iADSK, ADSK_YEAR_COL, iSF,
    If Not IsSNitemOK(ADSK_REGDAT_COL, SN_SF.Registered, SN_ADSK.Registered) Then _
        ErrSN iADSK, ADSK_REGDAT_COL, iSF, SFADSK_SNREGDAT_COL
    
    If Not SN_SF.ErrFlag Then IsSN_OK = True
    
End Function
Function IsSNitemOK(iMap, SN_SFitem, SN_ADSKitem) As Boolean
'
' - IsSNitemOK(iMap, SN_SFitem, SN_ADSKitem) проверки соотвествия отдельных арибутов SN
'   3.6.12

    IsSNitemOK = True
    If ADSK_RepMap(iMap) = "" Then Exit Function
    If LCase$(SN_SFitem) = LCase$(SN_ADSKitem) Then Exit Function
    IsSNitemOK = False
End Function
Sub ErrSN(iADSK, ColADSK, Optional iSF As Integer = 0, Optional ColSF As Integer = 0)
'
' - ErrSN(iADSK, ColADSK, iSF, ColSF) - сообщение о несоответствии и раскраска ADSKrep
'   4.6.12

    SN_SF.ErrFlag = True
    
    Sheets(ADSKrep).Select
    If ColADSK = ADSK_SN_COL Then _
        Range(Cells(iADSK, 1), Cells(iADSK, ADSK_HdrMapSize)).Interior.Color = rgbPink
    Cells(iADSK, ColADSK).Interior.Color = rgbRed
    
    Dim ValSF As String     '= значение поля в SF
    Dim ValADSK As String   '= значение поля в ADSK
    Dim Hdr As String       '= имя колонки
    
    Hdr = ADSK_HDR_Map(ColADSK)
    
    ValADSK = Cells(iADSK, ColADSK)
    If iSF <= 0 Or iSF > EOL_ADSKfrSF Or ColSF <= 0 Then
        ValSF = ""
    Else
        ValSF = Sheets(ADSKfrSF).Cells(iSF, ColSF)
    End If
    ErrMsg WARNING, "По отчету " & ADSKrep & " SN=" & SN_ADSK.SN & " несоответствие:" _
        & " в ADSK " & Hdr & "=" & ValADSK & ", а в SF =" & ValSF
End Sub
Function SNinSFatr(SN, iSF) As SNatr
'
' - SNinSFatr(SN, iSF) - возвращает атрибуты SN в SF по ADSKinSF
'   5.6.12

    SNinSFatr = ZeroSNatr()
        
    With Sheets(ADSKfrSF)
        Dim i As Integer
        For i = 2 To EOL_ADSKfrSF
            If SN = .Cells(i, SFADSK_SN_COL) Then
                SNinSFatr.Contr = .Cells(i, SFADSK_CONTRACT_COL)    ' SN найден
                SNinSFatr.AccN = .Cells(i, SFADSK_ACCN_COL)
                SNinSFatr.AccName = .Cells(i, SFADSK_ACCNAME_COL)
                SNinSFatr.C_Start = .Cells(i, SFADSK_CONTRSTARTDAT_COL)
                SNinSFatr.C_End = .Cells(i, SFADSK_CONTRENDDAT_COL)
                SNinSFatr.C_Status = .Cells(i, SFADSK_CONTRSTAT_COL)
                SNinSFatr.CM_Name = .Cells(i, SFADSK_CM_NAME_COL)
                SNinSFatr.CM_Mail = .Cells(i, SFADSK_CM_MAIL_COL)
                SNinSFatr.CM_Tel = .Cells(i, SFADSK_CM_TEL_COL)
                SNinSFatr.SN = SN
                SNinSFatr.Description = .Cells(i, SFADSK_DESCRIPTION_COL)
                SNinSFatr.Seats = .Cells(i, SFADSK_SEATS_COL)
                SNinSFatr.Status = .Cells(i, SFADSK_STATUS_SN_COL)
                SNinSFatr.Deployment = .Cells(i, SFADSK_DEPLOYMENT_COL)
                SNinSFatr.Sbs = .Cells(i, SFADSK_ISSBS_COL)
'                SNinSFatr.Release_Year = .Cells(i, )
                SNinSFatr.Registered = .Cells(i, SFADSK_SNREGDAT_COL)
                SNinSFatr.ErrFlag = False
                iSF = i
                Exit For
            End If
        Next i
    End With
End Function
Sub ContrADSKinSFatr(ContrADSK, iSF)
'
' - ContrADSKinSFatr(ContrADSK, iSF) - возвращает атрибуты котракта по ADSKinSF
'                                   выполняется только если SN в SF отсутствует
'   4.6.12

    SN_SF = ZeroSNatr()
    
    With Sheets(ADSKfrSF)
        Dim i As Integer
        For i = 2 To EOL_ADSKfrSF
            If ContrADSK = .Cells(i, SFADSK_CONTRACT_COL) Then      'ContrADSK есть?
                SN_SF.Contr = .Cells(i, SFADSK_CONTRACT_COL)
                SN_SF.AccN = .Cells(i, SFADSK_ACCN_COL)
                SN_SF.AccName = .Cells(i, SFADSK_ACCNAME_COL)
                SN_SF.C_Start = .Cells(i, SFADSK_CONTRSTARTDAT_COL)
                SN_SF.C_End = .Cells(i, SFADSK_CONTRENDDAT_COL)
                SN_SF.C_Status = .Cells(i, SFADSK_CONTRSTAT_COL)
                SN_SF.CM_Name = .Cells(i, SFADSK_CM_NAME_COL)
                SN_SF.ErrFlag = False
                iSF = i
                Exit For
            End If
        Next i
    End With
End Sub
Sub AccNinSFatr(AccN, iSF)
'
' - AccNinSFatr(AccN, iSF) - находит AccN в ADSKinSF и помещает в SN_SF
'   4.6.12
    
    SN_SF = ZeroSNatr()
    
    With Sheets(ADSKfrSF)
        Dim i As Integer
        For i = 2 To EOL_ADSKfrSF
            If AccN = .Cells(i, SFADSK_ACCN_COL) Then      ' AccN есть?
                SN_SF.AccN = .Cells(i, SFADSK_ACCN_COL)
                SN_SF.AccName = .Cells(i, SFADSK_ACCNAME_COL)
                SN_SF.ErrFlag = False
                iSF = i
                Exit For
            End If
        Next i
    End With
End Sub
Function ZeroSNatr() As SNatr
'
' - ZeroSNatr() - возвращаент инициализированную структуру SNatr
'   4.6.12
    
    ZeroSNatr.Contr = ""
    ZeroSNatr.AccN = ""
    ZeroSNatr.AccName = ""
    ZeroSNatr.C_Start = "1.1.1900"
    ZeroSNatr.C_End = "1.1.1900"
    ZeroSNatr.C_Status = ""
    ZeroSNatr.CM_Name = ""
    ZeroSNatr.CM_Mail = ""
    ZeroSNatr.CM_Tel = ""
    ZeroSNatr.SN = ""
    ZeroSNatr.Description = ""
    ZeroSNatr.Seats = 0
    ZeroSNatr.Status = ""
    ZeroSNatr.Deployment = ""
    ZeroSNatr.Sbs = "False"
    ZeroSNatr.Release_Year = ""
    ZeroSNatr.Registered = "1.1.1900"
    ZeroSNatr.ErrFlag = True        ' при инициализации предполагаем ОШИБКУ
End Function
Function SNinADSKatr(iADSK) As SNatr
'
' - SNinADSKatr(iADSK) - возвращает атрибуты SN по строке iADSK отчета ADSKrep
'   5.6.12

    SNinADSKatr.Contr = SNvalByMap(iADSK, ADSK_CONTR_COL)
    SNinADSKatr.AccN = SNvalByMap(iADSK, ADSK_ACCN_COL)
    SNinADSKatr.AccName = SNvalByMap(iADSK, ADSK_ACCNAME_COL)
    SNinADSKatr.C_Start = SNvalByMap(iADSK, ADSK_C_START_COL, "Date")
    SNinADSKatr.C_End = SNvalByMap(iADSK, ADSK_C_END_COL, "Date")
    SNinADSKatr.C_Status = SNvalByMap(iADSK, ADSK_C_STAT_COL)
    SNinADSKatr.CM_Name = SNvalByMap(iADSK, ADSK_CM_NAME_COL)
    If SNinADSKatr.CM_Name = "" Then _
        SNinADSKatr.CM_Name = SNvalByMap(iADSK, ADSK_CM_F_NAME_COL) _
            & " " & SNvalByMap(iADSK, ADSK_CM_L_NAME_COL)
    SNinADSKatr.CM_Mail = SNvalByMap(iADSK, ADSK_CM_MAIL_COL)
    SNinADSKatr.CM_Tel = SNvalByMap(iADSK, ADSK_CM_TEL_COL)
    SNinADSKatr.SN = SNvalByMap(iADSK, ADSK_SN_COL)
    SNinADSKatr.Description = SNvalByMap(iADSK, ADSK_SN_DESCR_COL)
    SNinADSKatr.Status = SNvalByMap(iADSK, ADSK_SN_STATUS_COL)
    SNinADSKatr.Deployment = SNvalByMap(iADSK, ADSK_DEPL_COL)
    SNinADSKatr.Seats = SNvalByMap(iADSK, ADSK_SEATS_COL, "Integer")
    SNinADSKatr.Sbs = SNvalByMap(iADSK, ADSK_SBS_COL, "Boolean")
    SNinADSKatr.Release_Year = SNvalByMap(iADSK, ADSK_YEAR_COL)
    SNinADSKatr.Registered = SNvalByMap(iADSK, ADSK_REGDAT_COL, "Date")
    SNinADSKatr.ErrFlag = False
End Function
Function SNvalByMap(iADSK, ColADSK, Optional ValType As String = "String") As String
'
' - SNvalByMap(iADSK, ColADSK, [ValType]) - возвращает значете поля типа ValType
'                   разбираются различные случаи умолчания и ошибки Val
'   5.6.12
    
    Dim RepIndex As String  '= значение mapping'а для поля в колонке ColADSK
    Dim RepMap  As Integer  '= значение mapping'а для поля в колонке ColADSK
    Dim Val As String       '= значение поля в отчете ADSK
    
    SNvalByMap = ""
    If ValType = "Boolean" Then SNvalByMap = "False"
    
    RepIndex = ADSK_RepMap(ColADSK)
    If ValType = "Date" Then SNvalByMap = "1.1.1900"
    If Not IsNumeric(RepIndex) Then Exit Function
    RepMap = RepIndex
    If RepMap <= 0 Or RepMap > ADSK_HdrMapSize Then Exit Function
    
    Val = Sheets(ADSKrep).Cells(iADSK, ADSK_RepMap(RepMap))
    
    SNvalByMap = Compressor(Val)
    If ValType = "String" Then
        Exit Function
    ElseIf ValType = "Integer" Then
        If Not (IsNumeric(Val) Or Val <= 0) Then Exit Function
    ElseIf ValType = "Date" And Not IsDate(Val) Then SNvalByMap = "1.1.1900"
    ElseIf ValType <> "Boolean" Then Exit Function
        If InStr(UCase$(Val), "T") <> 0 Then SNvalByMap = "True"
    End If
End Function


