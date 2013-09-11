Attribute VB_Name = "Declarations"
'-------------------------------------------------------------------
' Declarations - декларация структур, используемых в match 2.1
'
'  12.9.13

Option Explicit

'---- Каталог "Загрузки" или "Downloads" ----
Public Const DownloadDir = "C:\Users\Pavel_Khrapkin\Downloads\"
'============================ файлы DBs ===============================
Public DirDBs As String    ' каталог файлов DBs из 'match.xlsm'!We
' вспомогательный файл 'match_environment.xlsx' содержит DirDBs
Public Const F_match_environment = "C:\match_environment.xlsx"

Public Const F_MATCH = "match.xlsm"
Public Const F_1C = "1C.xlsx"
Public Const F_SFDC = "SFDC.xlsx"
Public Const F_ADSK = "ADSK.xlsm"
Public Const F_STOCK = "Stock.xlsx"
Public Const F_TMP = "W_TMP.xlsm"   '.xlsm потому, что содержит макросы

Public DB_MATCH As Workbook 'отчеты и таблицы match
Public DB_1C As Workbook    'отчеты 1C
Public DB_SFDC As Workbook  'отчеты Salesforce
Public DB_ADSK As Workbook  'отчеты Autodesk
Public DB_STOCK As Workbook 'отчеты по Складу и Заказам
Public DB_TMP As Workbook   'рабочий файл для записи "голубых" листов и WP

'==================== Процессы на листе Process ====================
Public Const Process = "Process"    'Каталог и контекст всех процессов
Public Const PROC_START = "<*>ProcStart"    ' начало Процесса
Public Const PROC_END = "<*>ProcEnd"        ' конец Процесса

Public Const PROC_TIME_COL = 1      'колонка 1 - дата и время выполнения Шага
Public Const PROC_STEP_COL = 2      'колонка Step - Шаг Процесса
Public Const PROC_PREVSTEP_COL = 3  'колонка PrevStep
Public Const PROC_STEPFILE_COL = 4  'колонка "В файле" - файл, где код Шага
Public Const PROC_STEPDONE_COL = 5  'колонка Done - Шаг выполнен
Public Const PROC_PAR1_COL = 6      'колонка Параметр 1 - имя Процесса
Public Const PROC_PAR2_COL = PROC_PAR1_COL + 1
Public Const PROC_NAME_COL = PROC_PAR1_COL ' имя Процесса в Par1
Public Const PROC_REP1_COL = 11     'колонка  - имя листа - Документ 1

Public Const PROCESS_NAME_COL = 7  'колонка - имя текущего Процесса
Public Const STEP_NAME_COL = 9     'колонка - имя текущего Шага
Public PublicProcName As String
Public PublicStepName As String

'-- загрузка и препроцессинг базы 1C.xlsm
Public Const REP_1C_P_LOAD = "Загрузка Платежей из 1С"
Public Const REP_1C_P_PAINT = "Раскраска Платежей 1С"
Public Const REP_1C_SFACCFIL = "Заполнение колонки 1 для Платежей 1С"

Public Const REP_1C_С_LOAD = "Загрузка Договоров из 1С"
Public Const REP_1C_С_PAINT = "Раскраска Договоров 1С"
Public Const REP_1C_A_LOAD = "Загрузка Справочника клиентов из 1С"

'-- загрузка и препроцессинг отчетов из SalesForce в базу SDFC.xlsm
Public Const REP_SF_LOAD = "Загрузка Платежей из Salesforce - SF"
Public Const REP_SFD_LOAD = "Загрузка Договоров из Salesforce - SFD"
    
'-- глобальные переменные, общие для всех отчетов
Public RepName As String    ' имя обрабатываемого отчета
Public Doing As String      ' строка в Application.StatusBar - что делает модуль
Public AllCol As Integer    ' Количество колонок в таблице отчета
Public Lines As Integer     ' количество строк текущего/нового отчета
Public LinesOld As Integer  ' количество строк старого отчета

Public Fruitful As Integer  ' счетчик полезных результатов

Public ExRespond As Boolean ' если False - завершение работы
Public Silent As Boolean    ' если True - не выводить сообщений об ошибках

'=============== База DB_MATCH - файл match.xlsm ==============
Public Const Header = "Header"              ' лист заголовков (Шапок) и Адаптеров
Public Const A_Dic = "A_Dic"                ' лист - Словарь Организаций

Public EOL_PaySheet     ' последняя строка отчета 1С по Платежам без пятки
Public EOL_DogSheet     ' последняя строка отчета 1С по Договорам без пятки
Public EOL_SF           ' последняя строка отчета по Платежам SF без пятки
Public EOL_SFD          ' последняя строка отчета по Договорам SFD без пятки
Public EOL_SFopp        ' последняя строка отчета по Проектам в SFopp без пятки
Public EOL_SFacc        ' последняя строка SFacc по Организациям без пятки
Public EOL_SForders     ' последняя строка отчета по Заказам SForders без пятки
Public EOL_ADSKfrSF     ' последняя строка отчета Autodesk из SF
Public EOL_Acc1C        ' последняя строка Справочника Организаций 1С
Public EOL_NewOpp       ' последняя строка листа новых Проектов
Public EOL_NewPay       ' последняя строка листа новых Платежей
Public EOL_NewContr     ' последняя строка листа новых Договоров
Public EOL_NewSN        ' последняя строка листа новых SN Autodesk
Public EOL_ContrLnk     ' последняя строка листа новых или обновленных Договоров
Public EOL_PaymentUpd   ' EOL листа изменений Платежей - связка Платежа с Договором
Public EOL_DIC          ' последняя строка листа Словаря Организаций
Public EOL_AdAcc        ' последняя строка листа новых Организаций для внесения в SF
Public EOL_AccntUpd     ' последняя строка листа новых связей Организаций Sf и 1С для внесения в SF
Public EOL_ADSKlnkPay   ' последняя строка листа новых связок Платежи-Контакт ADSK
Public EOL_SFlnkADSK    ' последняя строка листа связок Платежи-Контакт ADSK
Public EOL_ADSK         ' последняя строка листа отчета из ADSK.xlsx
Public EOL_Stock        ' последняя значащая строка листа по Складу
Public EOL_BTO          ' последняя значащая строка листа BTOlog
Public EOL_OrderList    ' последняя значащая строка листа Заказов

Public Const DATE_BULKY = "1.1.20"  ' дата окончания для Bulky Проектов
Public Const BIG = 77777            ' большое число для границ поиска

'------------- match TOC - Оглавление отчетов в базе данных ----------
Public Const TOC = "TOCmatch"           ' Оглавление листов всех файлов - баз данных
Public Const TOC_F_DIR_COL = 10         ' в TOCmatch записан Path файлов DBs - DirDBs

Public Const TOC_DATE_COL = 1           ' дата и время загрузки отчета
Public Const TOC_REPNAME_COL = 2        ' имя отчета в базе данных
Public Const TOC_EOL_COL = 3            ' EOL отчета без пятки
Public Const TOC_MYCOL_COL = 4          ' MyCol - число доп.колонок слева
Public Const TOC_RESLINES_COL = 5       ' число строк пятки ResLines
Public Const TOC_MADE_COL = 6           ' Made - завершенный шаг по листу
Public Const TOC_MAXDAYS_COL = 7        ' Периодичность - max дней с обновления отчета
Public Const TOC_REPFILE_COL = 8        ' имя файла, содержащего отчет
Public Const TOC_SHEETN_COL = 9         ' имя листа, содержащего отчет (окрашено)
Public Const TOC_STAMP_COL = 10         ' Штамп
Public Const TOC_STAMP_TYPE_COL = 11    ' Тип Штампа: строка (=) или подстрока (I)
Public Const TOC_STAMP_R_COL = 12       ' строка Штампа: (возможно, +EOL)
Public Const TOC_STAMP_C_COL = 13       ' колонка Штампа: (возможно, +MyCol)
Public Const TOC_CREATED_COL = 14       ' дата и время создания отчета
Public Const TOC_PAR_1_COL = 15         ' колонка Штампа Параметр 1
Public Const TOC_PAR_2_COL = 16         ' колонка Штампа Параметр 2
Public Const TOC_PAR_3_COL = 17         ' колонка Штампа Параметр 3
Public Const TOC_PAR_4_COL = 18         ' колонка Штампа Параметр 4
Public Const TOC_PAR_5_COL = 19         ' колонка Штампа Параметр 5
Public Const TOC_PAR_6_COL = 20         ' колонка Штампа Параметр 6
                        '--- при частичном обновлении отчета, а не при замене
Public Const TOC_FRDATEROW_COL = 22     ' строка отчета где FrDate
Public Const TOC_TODATEROW_COL = 23     ' строка отчета где ToDate
Public Const TOC_DATECOL_COL = 24       ' колонка отчета где даты
Public Const TOC_NEW_FRDATE_COL = 25    ' FrDate нового отчета
Public Const TOC_NEW_TODATE_COL = 26    ' ToDate нового отчета

Public Const TOC_PARCHECK_COL = TOC_PAR_1_COL   ' строка доп.Штампа
Public Const TOC_INSHEETN = TOC_PAR_2_COL       ' колонка - InSheetN
Public Const TOC_FORMNAME = TOC_PAR_2_COL       ' колонка - Имя Шаблона
Public Const TOC_FORMSUMMARY = TOC_PAR_3_COL    ' колонка - Имя Шаблона пятки
Public Const TOC_FRDATE_COL = TOC_PAR_4_COL     ' колонка- FrDate полного отчета
Public Const TOC_TODATE_COL = TOC_PAR_5_COL     ' колонка- ToDate полного отчета
Public Const TOC_REPLOADER_COL = TOC_PAR_6_COL  ' колонка- Loader отчета

Public Const TOCstart = 4       ' первая строка TOCmatch - вначале служебная часть
Public Const TOCrepLines = 8    ' первая стока Документов. после служебных строк
    
Type TOCmatch
    iTOC As Long        '=номер строки в таблице TOCmatch - Read ONLY!
    Dat As Date         '=Now   - дата и время загрузки отчета
    Name As String      'не изм.- имя отчета в базе данных
    EOL  As Long        '=изм.только MoveToMatch - EOL отчета без пятки
    MyCol As Long       '=изм.InsMyCol   - MyCol - число доп.колонок слева
    ResLines As Long    '=изм.InsSmmary  - число строк в пятке отчета после EOL
    Made As String      '=изм.каждый шаг - Made    - завершенный шаг по листу
    RepFile As String   'не изм.-  файл DB с отчетом
    SheetN As String    'не изм.-  имя листа, содержащего отчет
''''    Stamp As String     'не изм.-  строка - Штамп
''''    StampType As String 'не изм.-  Тип Штампа: строка (=) или подстрока (I)
''''    StampR As Long      'не изм.-  строка Штампа: (+EOL)
''''    StampC As Long      'не изм.-  колонка Штампа: (+MyCol)
    CreateDat As Date   '=изм.только MoveToMatch и NewSheet -  дата и время создания отчета
''''    ParChech As String  'не изм.-   <>"" - след.строка - доп.Штамп
''''    Loader As String    'не изм.-   Loader отчета
    FormName As String  'не изм.-   Название формы - Шаблона
End Type

Public RepTOC As TOCmatch   ' структура TOCmatch
''''''''''''Public RepMatch As TOCmatch ' структура TOCmatch для отчета в match.xlsm
''''''''''''Public Rep1C As TOCmatch    ' структура TOCmatch для отчета в 1С.xlsm
''''''''''''Public RepSF As TOCmatch    ' структура TOCmatch для отчета в SFDC.xlsm
''''''''''''Public RepADSK As TOCmatch  ' структура TOCmatch для отчета в ADSK.xlsm
''''''''''''Public RepStock As TOCmatch ' структура TOCmatch для отчета в Stock.xlsm

Public Const REP_LOADED = "Loaded"            ' MoveToMatch: отчет загружен в файл DB

'=============== База DB_SFDC - файл SFDC.xlsm ==============
Public Const SF = "SF"              ' лист отчета по Платежам
Public Const SFD = "SFD"            ' лист отчета по Договорам
Public Const SFacc = "SFacc"        ' лист - список Организаций
Public Const SFcont = "SFcont"      ' лист отчета по Контактам
Public Const SFopp = "SFopp"        ' лист отчета по Проектам
Public Const SForders = "SForders"  ' лист отчета по Заказам
Public Const ADSKfrSF = "ADSKfrSF"  ' лист отчета по Autodesk

Public Const SFresLines = 6 'размер пятки отчетов SalesForce

'------------- SF - отчет Saleforce по Платежам --------------------
Public Const SF_MATCH_COL = 1      ' колонка-Match - номер строки Платежа 1С
Public Const SF_COD_COL = 2        ' колонка-код Платежа в SF
Public Const SF_ACC1C_COL = 3      ' колонка-"Имя плательщика в 1С"
Public Const SF_INV_COL = 8        ' колонка-"Номер счета"
Public Const SF_SAIL_COL = 9       ' колонка- Сотрудник, ответственный за проект
Public Const SF_PAYID_COL = 18     ' колонка- Id Платежа в SF
    
'------------- SFD - отчет Saleforce по Договорам ------------------
Public Const SFD_COD_COL = 2       ' колонка-код "Осн/Дог" в SFD
Public Const SDF_CONTRN_COL = 3    ' колонка Номер договора в SF
Public Const SFD_ACC1C_COL = 6     ' колонка "Имя Организации в 1С"
Public Const SFD_DATESTART_COL = 8 ' колонка "Дата начала Договора"
Public Const SFD_DATEEND_COL = 9   ' колонка "Дата окончания Договора"
Public Const SFD_SCAN_COL = 10     ' колонка "Отсканировано"
Public Const SFD_OWNER_COL = 12    ' колонка "Отв."
Public Const SFD_STATUS_COL = 13   ' колонка "Статус"
Public Const SFD_CONTRID_COL = 15  ' колонка "Код договора"
Public Const SFD_OPPN_COL = 17     ' колонка OppN в SFD
Public Const SFD_OPPID_COL = 18    ' колонка OppId
Public Const SFD_VALUE_COL = 20    ' колонка Сумма Договора

'--------------- SFopp Проекты ----------------------------------------
Public Const SFOPP_OPPID_COL = 1        ' колонка "Код Проекта"
Public Const SFOPP_OPPN_COL = 2         ' колонка "Номер Проекта"
Public Const SFOPP_OPPNAME_COL = 3      ' колонка "Имя Проекта"
Public Const SFOPP_ACC1C_COL = 4        ' колонка Оргинизация 1С листа SFopp
Public Const SFOPP_PROBABILITY_COL = 5  ' колонка "Вероятность (%)"
Public Const SFOPP_CRDAT_COL = 8        ' колонка Дата создания Проекта SFopp
Public Const SFOPP_CLOSEDATE_COL = 9    ' колонка "Дата закрытия" Проекта
Public Const SFOPP_SALE_COL = 10        ' колонка "Ответственный за Проект"
Public Const SFOPP_TO_PAY_CUR_COL = 14  ' колонка "Еще не оплачено Валюта" листа SFopp
Public Const SFOPP_TO_PAY_VAL_COL = 15  ' колонка "Еще не оплачено" листа SFopp
Public Const SFOPP_LINE_COL = 17        ' колонка "Направление деятельности"
Public Const SFOPP_TYP_COL = 18         ' колонка "Тема&Вид деятельности" Проекта

'--------------- SFacc Организации ------------------------------------
Public Const SFACC_ACCNAME_COL = 1      ' колонка "Имя oрганизации" в SF
Public Const SFACC_ACC1C_COL = 2        ' колонка "Имя Организации в 1С"
Public Const SFACC_IDACC_COL = 3        ' колонка "Код организации"
Public Const SFACC_STREET_COL = 4       ' колонка "Улица для счета"
Public Const SFACC_CITY_COL = 5         ' колонка "Город для счета"
Public Const SFACC_STATE_COL = 6        ' колонка "Область для счета"
Public Const SFACC_INDEX_COL = 7        ' колонка "Индекс для счета"
Public Const SFACC_COUNTRY_COL = 8      ' колонка "Страна для счета"

'--------------- SFADSK - отчет по Autodesk из SF ------------------------
Public Const SFADSK_CONTRACT_COL = 19   ' колонка "Номер Контракта Autodesk"
Public Const SFADSK_CONTRID_COL = 2     ' колонка "Koд Контракта Autodesk" = Id SF
Public Const SFADSK_CONTRSTAT_COL = 3   ' колонка "Status" - Статус Контракта Autodesk
Public Const SFADSK_SN_COL = 4          ' колонка "SN продукта Autodesk"
Public Const SFADSK_DESCRIPTION_COL = 5 ' колонка "Имя Продукта Autodesk"
Public Const SFADSK_SEATS_COL = 6       ' колонка "Seats" - мест
Public Const SFADSK_STATUS_SN_COL = 7   ' колонка "Статус SN"
Public Const SFADSK_ACCN_COL = 8        ' колонка "Account_#"
Public Const SFADSK_ACCNAME_COL = 9     ' колонка "Account_Name ADSK"
Public Const SFADSK_CONTRSTARTDAT_COL = 11 ' колонка "Contract Start Date"
Public Const SFADSK_CONTRENDDAT_COL = 12   ' колонка "Contract End Date"
Public Const SFADSK_SALE_COL = 15       ' колонка "Продавец"
Public Const SFADSK_ISSBS_COL = 20        ' колонка "Есть подписка"
Public Const SFADSK_SNREGDAT_COL = 21   ' колонка "Дата активации"
Public Const SFADSK_ACC1C_COL = 23      ' колонка "Имя oрганизации в 1С"
Public Const SFADSK_CM_NAME_COL = 24    ' колонка "Контакт СМ: Полное имя"
Public Const SFADSK_CM_MAIL_COL = 25    ' колонка "CM мейл"
Public Const SFADSK_CM_TEL_COL = 26     ' колонка "CM телефон"
Public Const SFADSK_DEPLOYMENT_COL = 27 ' колонка "Deployment"

Public Const SFADSK_SN_REGISTERED = "Registered"   ' "Статус SN" Registered

'- - - - - - - - SF_PA = "SF_P_ADSK_LINK" - - - - - - - - - - - - - -
Public Const SF_PA = "SF_PA"        ' лист Связей Платеж-Контракт ADSK

Public Const SFPA_PAYID_COL = 2     ' колонка "Платеж: Код Платеж"
Public Const SFPA_ADSKID_COL = 3    ' колонка "Код Контракта ADSK"

'. . . .  структура листа типа SNatr  - лист, загруженный из ADSK.xlsx . . . .
'----- Оглавление базы ADSK.xlsx --------------------
Public Const TOC_ADSK = "TOC_ADSK"
Public Const TOCADSK_REPNAME_COL = 3      'поле - Имя/тип отчета
Public Const TOCADSK_REPRANGE_COL = 5     'поле - Лист (Range)

Public ADSKrep                      ' имя отчета из ADSK.xlsx
Public Const ADSK_HdrMapSize = 20           ' размер ADSK_RepMap
Public ADSK_RepMap(1 To ADSK_HdrMapSize)    ' структура полей отчета ADSKrep
Public ADSK_HDR_Map(1 To ADSK_HdrMapSize)   ' заголовки полей отчета из ADSK

Public SN_SF As SNatr               ' SN со всеми атрибутами из SF
Public SN_ADSK As SNatr             ' SN со всеми атрибутами - строка из ADSK.xlsx

Public Const ADSK_CONTR_COL = 1     ' Колонка "ContrADSK" - Контракт Autodesk / Agreement
Public Const ADSK_ACCN_COL = 2      ' Колонка "AccN" - CSN на PartnerCenter
Public Const ADSK_ACCNAME_COL = 3   ' Колонка "AccName" - Имя клиента на PartnerCenter
Public Const ADSK_C_START_COL = 4   ' Колонка "ContrStartDate" - Дата начала Контракта ADSK
Public Const ADSK_C_END_COL = 5     ' Колонка "ContractEndDate" - Дата окочания Контракта             ' Дата окончания Контракта ADSK
Public Const ADSK_C_STAT_COL = 6    ' Колонка "Contr_Status" - Статус Контракта ADSK
Public Const ADSK_CM_NAME_COL = 8   ' Колонка Имя СМ - контакта по SN
Public Const ADSK_CM_L_NAME_COL = 9 ' Колонка Имя СМ - ФАМИЛИЯ
Public Const ADSK_CM_F_NAME_COL = 10 ' Колонка Имя СМ - ИМЯ
Public Const ADSK_CM_MAIL_COL = 11  ' Колонка email CM
Public Const ADSK_CM_TEL_COL = 12   ' Колонка телефон СМ
Public Const ADSK_SN_COL = 13       ' Колонка "SN"
Public Const ADSK_SN_DESCR_COL = 14 ' Колонка описание продукта Autodesk
Public Const ADSK_SN_STATUS_COL = 15 ' Колонка Статус SN: Registered, Upgrade...
Public Const ADSK_DEPL_COL = 16     ' Колонка "Deployment" лицензия Сетевая или Локальная
Public Const ADSK_SEATS_COL = 17    ' Колонка Число лицензий/посадочных мест
Public Const ADSK_SBS_COL = 18      ' Колонка "Subscriptions" = Sbs - есть ли подписка              ' TRUE если подписка активна
Public Const ADSK_YEAR_COL = 19     ' Колонка "Release_Year" - Год релиза продукта ADSK
Public Const ADSK_REGDAT_COL = 20   ' Колонка Дата регистрации SN

'. . . .  ADSK_C - лист Контрактов ADSK из PartnerCenter . . . .
Public Const ADSK_C = "ADSK_C"      ' лист Контрактов ADSK

Public Const ADSK_C_STAMP = "Contract #"

Public Const ADSK_C_CONTR_COL = 2   ' колонка "Contract #"
Public Const ADSK_C_ENDDATE_COL = 3 ' колонка "Contract End Date"
Public Const ADSK_C_ACCN_COL = 11   ' колонка "End Customer CSN"

'-------------- Платежи - отчет из 1С ---------------------------------
Public Const PAY_SHEET = "Платежи"  ' лист Платежей

''' Public Const PAY_RESLINES = 3       ' кол-во строк пятки отчета по Платежам
''' Public Const PAY_MYCOLS = 5         ' количество моих колонок слева

Public Const PAYISACC_COL = 1       ' колонка =1, если Организация есть в SF
Public Const PAYINSF_COL = 2        ' колонка =1, если Платеж есть в SF
Public Const PAYIDSF_COL = 3        ' колонка - Id Платежа в SF
Public Const PAYADSK_COL = 4        ' колонка - связи Платежа с ADSK в SF
Public Const PAYCODE_COL = 5        ' колонка кода Платежа
'----
Public Const PAYDOC_COL = 6         ' колонка "Плат.док."
Public Const PAYDATE_COL = 7        ' колонка "Дата прих.денег"
Public Const PAYINVOICE_COL = 8     ' колонка "Счет"
Public Const PAYACC_COL = 9         ' колонка "Клиент"
Public Const PAYINVVAL_COL = 10     ' колонка "Сумма счета в руб"
Public Const PAYRUB_COL = 18        ' колонка "Итог руб"
Public Const PAYGOOD_COL = 19       ' колонка Товар
Public Const PAYFIRM_COL = 20       ' колонка Фирма
Public Const PAYSALE_COL = 22       ' колонка Продавец
Public Const PAYDOGOVOR_COL = 25    ' колонка Договор
Public Const PAYOSNDOGOVOR_COL = 26 ' колонка Осн.Договор
Public Const PAYDOGVAL_COL = 27     ' колонка Сумма договора
Public Const PAYDOGCUR_COL = 28     ' колонка Валюта договора
Public Const PAYINV_COL = 30        ' колонка CX-XXX - Счет без даты
Public Const PAYGOODTYPE_COL = 31   ' колонка Тип товара

'''Public Const Stamp1Cpay1 = "Плат. док.", Stamp1Cpay2 = "Дата прих. денег"
'-------------- Договоры - отчет из 1С ---------------------------------
Public Const DOG_SHEET = "Договоры" ' лист Договоров для модуля ConctAnalitics

Public Const DOGRES = 7             ' кол-во строк пятки отчета по Договорам

Public Const DOGISACC_COL = 1      ' колонка =1, если Организация есть в SF
Public Const DOGSFSTAT_COL = 2     ' колонка статуса Договора в SF
Public Const DOGIDSF_COL = 3       ' колонка Id Договора в SF
Public Const DOGPAID1C_COL = 5     ' колонка Договора1C "Оплачено"
Public Const DOGISINV1C_COL = 6    ' колонка Договора1C "по счетам"
Public Const DOG1CSCAN_COL = 7     ' колонка "Отсканировано"
Public Const DOGCOD_COL = 8        ' колонка кода Договора вида "Осн/Договр"
'----
Public Const DOG1C_COL = 10        ' колонка "Номер договора" в отчете 1С
Public Const DOG1CDAT_COL = 12     ' колонка Дата подписания
Public Const DOG1CFIRM_COL = 13    ' колонка Фирма
Public Const DOG1CSALE_COL = 14    ' колонка Продавец
Public Const DOG1CACC_COL = 15     ' колонка Заказчик
Public Const VENDOR1C_COL = 16     ' колонка Поставщик листа Договоров 1С
Public Const DOGTYPE_COL = 17      ' колонка Тип Договора
Public Const DOG1C_MAINDOG_COL = 18 ' колонка "Основной договор"
Public Const DOGSUM1C_COL = 20     ' колонка Сумма (план)
Public Const DOGINVC_COL = 21      ' колонка Сумма по счетам
Public Const DOGPAID_COL = 22      ' колонка Сумма по платежам
Public Const DOGCUR1C_COL = 23     ' колонка Валюта Договора

Public Const DOG_STAT_DRAFT = "Черновик"
Public Const DOG_STAT_OPEN = "Открыт"
Public Const DOG_STAT_CLOSED = "Закрыт"
Public Const DOG_STAT_CANCEL = "Не состоялся"
Public Const DOG_STAT_NOTINSF = "Нет в SF"

'-------------- Справочник Организаций - отчет из 1С -------------------
Public Const Acc1C = "Список клиентов 1C"   ' Справочник клиентов 1С

Public Const ACC1C_RES = 3      ' пятка справочника организаций 1С

Public Const A1C_ISACC_COL = 1  ' колонка =1, если Организация есть в SF
'---
Public Const A1C_NAME_COL = 5   ' колонка "Название фирмы" в 1С
Public Const A1C_CON_COL = 6    ' колонка "контакт" в 1С
Public Const A1C_TEL_COL = 7    ' колонка "Телефон"
Public Const A1C_ADR_COL = 8    ' колонка "Адрес" в 1С
Public Const A1C_FACTADR_COL = 9 ' колонка "Факт. адрес" в 1С
Public Const A1C_INN_COL = 10   ' колонка "ИНН" в 1С
Public Const A1C_INVOICE_COL = 17  ' колонка "Список счетов с оплатой" в 1С
Public Const A1C_GOOD_COL = 18  ' колонка "Товар последнего счета" в 1С

'~~~~~~~~~~~~~~~~~~~~~~~~ Склад ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Const STOCK_SHEET = "Склад"  ' Лист проводок по Складу
Public Const STOCK_STAMP = "Наименование продукта"
Public Const STOCK_MYCOLS = 6       ' количество моих колонок слева

Public Const STOCK_NEWSN_COL = 1    ' колонка "SN"
Public Const STOCK_CONTRADSK_COL = 2 ' колонка "ContrADSK"
Public Const STOCK_SF_SN_COL = 3    ' колонка "SF_SN"
Public Const STOCK_GOOD_COL = 4     ' колонка "Товар"
Public Const STOCK_IDSF_COL = 5     ' колонка "IdSF"
Public Const STOCK_IDSFORDER_COL = 6 ' колонка "IdSForder"
Public Const STOCK_INVOICE_COL = 7  ' колонка "Счет"
Public Const STOCK_ACC_COL = 8      ' колонка "Acc"
Public Const STOCK_CLIENT_COL = 9   ' колонка "Клиент"
Public Const STOCK_PRODUCT_COL = 10  ' колонка "Наименование продукта"
Public Const STOCK_SN_COL = 12      ' колонка "Серийный №"
Public Const STOCK_DATE_COL = 13    ' колонка "Дата"
Public Const STOCK_OLDSN_COL = 15   ' колонка "Старый S/N"

'~ ~ ~ ~ ~ ~ ~ ~ ~ Письма БТО не проведенные по Складу ~ ~ ~ ~ ~ ~ ~ ~ ~
Public Const BTOmails = "BTOmails"  ' Исходные письма БТО из файла BTOmails.txt
Public Const BTO_SHEET = "BTOlog"   ' Письма БТО без проводки по Складу

Public Const BTO_DATE_COL = 1       ' колонка "Дата письма"
Public Const BTO_INVCSD_COL = 2     ' колонка "Заказ/Счет CSD"
'''Public Const BTO_INV1C_COL = 3      ' колонка "Счет 1С" с датой счета
'''Public Const BTO_SALE_COL = 4       ' колонка "Продавец"
'''Public Const BTO_CLIENT_COL = 5     ' колонка "Заказчик"
Public Const BTO_SN_COL = 7         ' колонка "SN"
''Public Const BTO_GOOD_COL = 7       ' колонка "Товар ADSK"
''Public Const BTO_SEATS_COL = 8      ' колонка "Мест"
Public Const BTO_UNIQUE_COL = 11    ' колонка "UNIQUE"
''Public Const BTO_MAIL_COL = 9       ' колонка "Письмо"

'############## Заказы у Дистрибуторов или Поставщиков #################
Public Const ORDER_SHEET = "Заказы" ' Лист Заказов

Public Const OL_IDSF_COL = 1        ' колонка "IdSF"
Public Const OL_GOOD_COL = 14       ' колонка "TMPgoodType" - Категория товара в Заказе
Public Const OL_INV1C_COL = 15      ' колонка "TMPinv1C" - Счет 1С в Заказе
Public Const OL_CUSTOMER_COL = 16   ' колонка "TMPcustomer" - Клиент в Заказе
Public Const OL_SALES_COL = 17      ' колонка "TMPsalesRep" - Продавец в Заказе
Public Const OL_ORDERNUM_COL = 18   ' колонка "OrderNumber" - Продавец в Заказе
Public Const OL_CSDINVDAT_COL = 20  ' колонка "CSDinvDate" - дата Счета CSD в Заказе

'................. Лист новых Организаций A_Acc ..................
Public Const A_Acc = "A_Acc"            ' лист новых Организаций

Public Const ADACC_NAME_COL = 1         ' колонка "AccountName"
Public Const ADACC_1CNAME_COL = 2       ' колонка "Acc1CName"
Public Const ADACC_INN_COL = 4          ' колонка "INN"
Public Const ADACC_STREET_COL = 5       ' колонка "AccStreet"
Public Const ADACC_CITY_COL = 6         ' колонка "AccCity"
Public Const ADACC_INDEX_COL = 7        ' колонка "AccIndex"
Public Const ADACC_STATE_COL = 8        ' колонка "AccState"
Public Const ADACC_COUNTRY_COL = 9      ' колонка "AccCountry"
Public Const ADACC_TEL_COL = 14         ' колонка "Телефон"
Public Const ADACC_FACTSTREET_COL = 17  ' колонка "FactStreet"
Public Const ADACC_FACTCITY_COL = 18    ' колонка "FactCity"
Public Const ADACC_FACTINDEX_COL = 19   ' колонка "FactIndex"
Public Const ADACC_FACTSTATE_COL = 20   ' колонка "FactState"
Public Const ADACC_FACTCOUNTRY_COL = 21 ' колонка "FactCountry"
Public Const ADACC_CONTACT1C_COL = 22   ' колонка "Contact1C" - не загружать в DL!

'........ Лист новых связей Организаций существующих в SF и 1С ...........
Public Const AccntUpd = "AccntUpd"      ' лист новых связей Организаций SF и 1С

Public Const ACCUPD_SFID_COL = 1        ' колонка "SFaccId"
Public Const ACCUPD_1CNAME_COL = 2      ' колонка "Acc1C"

'................ Лист новых Платежей NewPayment ..........................
Public Const NEW_PAYMENT = "NewPayment" ' лист новых Платежей

Public Const NEWPAY_DOC_COL = 1         ' Doc - Имя платежного документа
Public Const NEWPAY_DATE_COL = 2        ' Date - Дата платежа
Public Const NEWPAY_INV_COL = 3         ' Invoice - Счет

'................ Лист новых Договоров NewContract ..........................
Public Const NEW_CONTRACT = "NewPayment" ' лист новых Договоров

Public Const NEWDOG_DOGOVOR_COL = 1     ' Dogovor - Имя Договора
Public Const NEWDOG_DATE_COL = 2        ' Date - Дата подписания
Public Const NEWDOG_WE_COL = 3          ' We - наша фирма
Public Const NEWDOG_IDWE_COL = 4        ' IdWe
Public Const NEWDOG_SALESMAN_COL = 5    ' Salesman - Продавец
Public Const NEWDOG_IDOWNER_COL = 6     ' IdOwner - Id Продавца
Public Const NEWDOG_CLIENT_COL = 7      ' Client - Заказчик
Public Const NEWDOG_IDACC_COL = 8       ' IdAcc - Id Заказчика
Public Const NEWDOG_TYPE_COL = 9        ' Type - ContractType - Тип договора
Public Const NEWDOG_MAINDOG_COL = 10    ' MainContract - Основной договор
Public Const NEWDOG_CUSTCONTR_COL = 11  ' CustomerContractNumber
Public Const NEWDOG_VALUE_COL = 12      ' Value - Сумма(план)
Public Const NEWDOG_CURRENCY_COL = 13   ' Currency - Валюта
Public Const NEWDOG_CURRENCYCOD_COL = 14 ' CurrencyCod - код ISO валюты
Public Const NEWDOG_RATE_COL = 15       ' CurrencyExchangeRate - курс валюты
Public Const NEWDOG_RATEDATE_COL = 16   ' курс валюты на дату
Public Const NEWDOG_ACCSTREET_COL = 17  ' улица
Public Const NEWDOG_ACCCITY_COL = 18    ' город
Public Const NEWDOG_ACCINDEX_COL = 19   ' индекс
Public Const NEWDOG_ACCSTATE_COL = 20   ' область
Public Const NEWDOG_ACCCOUNTRY_COL = 21 ' страна

Public Const DOG_UPDATE = "ContractUpdate" ' лист изменений (связок с Opp) Договоров
'................ Лист новых Проектов NewOpp ..........................
Public Const NEW_OPP = "NewOpp"         '"голубой" лист новых Проектов

Public Const NEWOPP_ACCID_COL = 1       ' AccId
Public Const NEWOPP_OPPNAME_COL = 2     ' OppName
Public Const NEWOPP_CLOSDATE_COL = 3    ' CloseDate
Public Const NEWOPP_OWNERID_COL = 4     ' OwnerId
Public Const NEWOPP_VALUE_COL = 5       ' Value
Public Const NEWOPP_ISOCUR_COL = 6      ' ISOcur
Public Const NEWOPP_TYPOPP_COL = 7      ' TypOpp                    Type = Тип
Public Const NEWOPP_STAGENAME_COL = 8   ' StageName
Public Const NEWOPP_LINE_COL = 9        ' OpportunityLineOfActivity Направление деятельности
Public Const NEWOPP_KIND_COL = 10       ' OpportunityActivityKind   Тема&Вид деятельности
Public Const NEWOPP_BUDDY_COL = 11      ' Buddy
Public Const NEWOPP_OPPUNIQ_COL = 12    ' OppUniq

'................ Лист новых Серийных номеров Autodesk .................
Public Const NewSN = "NewSN"            ' Лист NewSN

Public Const NEWSN_SN_COL = 1           ' SN
Public Const NEWSN_CONTRADSK_COL = 2    ' ContrADSK
Public Const NEWSN_ACCN_COL = 3         ' ФссТ

'................ Лист обработки Серийных номеров Autodesk .................
Public iActive As Long '- номер строки в SNactive
Public iUpdate As Long '- номер строки в SNupdate
Public nActive As Long '= счетчик Registered
Public nUpdate As Long '= счетчик Update
Public nToClean As Long '=счетчик "Разбираться вручную"

'.............. Лист SNtmp - SN из  License Inquiry PCв Autodesk ...........
Public Const SNTMP_SN_COL = 1           ' SN
Public Const SNTMP_ACTDATE_COL = 9      ' Most Recent Activation Date
Public Const SNTMP_STATUS_COL = 12      ' Статус SN
Public Const SNTMP_STATUS_REGISTERED = "Registered"
    
'----------------------------------------------------------------------------
Public Const WP = "WP"                  ' лист WP

'****************** Лист We - Словари, таблицы *************************
Public Const We = "We"  ' Лист: наши юр.лица, Продавцы, коды валют, типы товаров

Public Const WE_GOODS_COL = 1           ' Продукты (Goods)
Public Const WE_GOODS_LINCOL = 3        ' Направление деятельности
Public Const WE_GOODS_KINDCOL = 4       ' Тема & вид деятельности
Public Const WE_GOODS_ISSBSCOL = 8      ' Товар - Есть подписка
Public Const WE_GOODS_NOSBSCOL = 9      ' Товар - Нет подписки

Public Const WE_GOODS_ADSK = "Autodesk" ' Товар - Autodesk

'########################## лист лидов MS CSIT ###########################
Public Const CSIT_MS = "CSIT_MS"            ' лист Лидов CSIT по Microsoft
Public EOL_CSIT_MS                          ' последняя строка листа CSIT_MS

Public Const CSIT_MS_IDSF_COL = 1           ' Колонка "IdSF" - Id соответсветвствующих
                                            '                    Организаций в SF
Public Const CSIT_MS_NAME_COL = 2           ' Колонка "Название" - имя организации
Public Const CSIT_MS_ADDR_COL = 5           ' Колонка "Адрес"
Public Const CSIT_MS_STAMP = "Первое Независимое Рейтинговое Агентство"

'!!!!!!!!!!!!!!!!!!!!!!!!!!!! Типы данных !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Type PostAddr               'Атрибуты почтового адреса
    PostIndex As String         ' Индекс
    City As String              ' Город
    State As String             ' Область/Район/Край
    Street As String            ' Улица, дом, здание
    Country As String           ' Страна
    ErrFlag As Boolean          '= TRUE - данные адреса неконсистентны
End Type

Type SNatr                  'Атрибуты SN Autodesk
    sN      As String           ' Серийный номер
    Contr   As String           ' Контракт Autodesk / Agreement
    AccN    As String           ' Customer Account Number, CSN на PartnerCenter
    AccName As String           ' Имя клиента на PartnerCenter
    C_Start As Date             ' Дата начала Контракта ADSK
    C_End   As Date             ' Дата окончания Контракта ADSK
    C_Status As String          ' Статус Контракта: Active, Expired, Inactive
    CM_Name As String           ' Имя СМ - контакта по SN
    CM_Mail As String           ' email CM
    CM_Tel  As String           ' телефон СМ
    Description As String       ' описание продукта Autodesk
    Seats As Integer            ' Число лицензий/посадочных мест
    Status As String            ' Статус SN: Registered, Upgrade...
    Deployment As String        ' Тип лицензии Сетевая или Локальная
    Sbs As Boolean              ' TRUE если подписка активна
    Release_Year As String      ' Год релиза продукта ADSK
    Registered As Date          ' Дата регистрации SN
    ErrFlag As Boolean          '= TRUE - данные/атрибуты SN неконсистентны
End Type

'<><><><><><><> Коды сообщений об ошибках <><><><><><><><><><>
Public Const FATAL_ERR = 1  ' вывести сообщение в Log лист, на экран и остановиться
Public Const TYPE_ERR = 2   ' вывести сообщение в Log лист и на экран
Public Const WARNING = 3    ' вывести предупреждение только в Log лист


