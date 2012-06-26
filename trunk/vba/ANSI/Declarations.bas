Attribute VB_Name = "Declarations"
'-------------------------------------------------------------------
' Declarations - декларация структур листов Match
'
'   А.Пасс & П.Храпкин
'   17.6.12 - А.Пасс
'   18.6.12 - ПХ merge с версией АП от 17/6
'   22.6.12 - А.Пасс definitions for fax support

Option Explicit
    
'------------- SFD -------------------------------------------------
Public Const SFD_COD_COL = 2       ' колонка-код "Осн/Дог" в SFD
Public Const SDF_CONTRN_COL = 3    ' колонка Номер договора в SF
Public Const SFD_ACC1C_COL = 6     ' колонка "Имя Организации в 1С"
Public Const SFD_DATESTART_COL = 8 ' колонка "Дата начала Договора"
Public Const SFD_DATEEND_COL = 9   ' колонка "Дата окончания Договора"
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
Public Const SFACC_INN_COL = 12         ' колонка "INN"
Public Const SFACC_TEL_COL = 13         ' колонка "TEL"
Public Const SFACC_DELSTREET_COL = 14   ' колонка "Улица для доставки"
Public Const SFACC_DELCITY_COL = 15     ' колонка "Город для доставки"
Public Const SFACC_DELSTATE_COL = 16    ' колонка "Область для доставки"
Public Const SFACC_DELINDEX_COL = 17    ' колонка "Индекс для доставки"
Public Const SFACC_DELCOUNTRY_COL = 18  ' колонка "Страна для доставки"

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

Public Const ADSKfrSFstamp = "Match ADSK from SF"   ' штамп - имя Отчета SF в пятке

Public Const SFADSK_SN_REGISTERED = "Registered"   ' "Статус SN" Registered

'- - - - - - - - SF_PA = "SF_P_ADSK_LINK" - - - - - - - - - - - - - -
Public Const SF_PA = "SF_PA"        ' лист Связей Платеж-Контракт ADSK

Public Const SFPA_PAYID_COL = 2     ' колонка "Платеж: Код Платеж"
Public Const SFPA_ADSKID_COL = 3    ' колонка "Код Контракта ADSK"

Public Const SFpaRepName = "Match: Связка Платеж-Контракт Autodesk"

'. . . .  структура листа типа SNatr  - лист, загруженный из ADSK.xlsx . . . .
'----- Оглавление базы ADSK.xlsx --------------------
Public Const TOC_ADSK = "TOC_ADSK"
Public Const TOC_REPNAME_COL = 3      'поле - Имя/тип отчета
Public Const TOC_REPRANGE_COL = 5     'поле - Лист (Range)

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
Public Const PAY_SHEET = 1          ' лист Платежей

Public Const PAY_RESLINES = 3       ' кол-во строк пятки отчета по Платежам

Public Const PAYINSF_COL = 4        ' колонка =1, если Платеж есть в SF
Public Const PAYCODE_COL = 5        ' колонка кода Платежа
Public Const PAYDOC_COL = 6         ' колонка "Плат.док."
Public Const PAYDATE_COL = 7        ' колонка "Дата прих.денег"
Public Const PAYINVOICE_COL = 8     ' колонка "Счет"
Public Const PAYACC_COL = 9         ' колонка "Клиент"
Public Const PAYRUB_COL = 18        ' колонка "Итог руб"
Public Const PAYGOOD_COL = 19       ' колонка Товар
Public Const PAYSALE_COL = 22       ' колонка Продавец
Public Const PAYDOGOVOR_COL = 25    ' колонка Договор
Public Const PAYOSNDOGOVOR_COL = 26 ' колонка Осн.Договор

Public Const Stamp1Cpay1 = "Плат. док.", Stamp1Cpay2 = "Дата прих. денег"
'-------------- Договоры - отчет из 1С ---------------------------------
Public Const DOG_SHEET = 3         ' лист Договоров для модуля ConctAnalitics

Public Const DOGRES = 7             ' кол-во строк пятки отчета по Договорам

Public Const DOGSFSTAT_COL = 2     ' колонка статуса Договора в SF
Public Const DOGPIAID1C_COL = 5    ' колонка Договора1C "Оплачено"
Public Const DOG1CSCAN_COL = 7     ' колонка "Отсканировано"
Public Const DOGCOD_COL = 8        ' колонка кода Договора вида "Осн/Договр"
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

Public Const Stamp1Cdog1 = "Номер договора", Stamp1Cdog2 = "Дата подписания"
Public Const Stamp1Cacc1 = "Название фирмы", Stamp1Cacc2 = "Контакт"

'-------------- Справочник Организаций - отчет из 1С -------------------
Public Const Acc1C = "Список клиентов 1C"   ' Справочник клиентов 1С

Public Const ACC1C_RES = 3      ' пятка справочника организаций 1С

Public Const A1C_NAME_COL = 3   ' колонка "Название фирмы" в 1С
Public Const A1C_CON_COL = 4    ' колонка "контакт" в 1С
Public Const A1C_TEL_COL = 5    ' колонка "Телефон"
Public Const A1C_ADR_COL = 6    ' колонка "Адрес" в 1С
Public Const A1C_FACTADR_COL = 7 ' колонка "Факт. адрес" в 1С
Public Const A1C_INN_COL = 8    ' колонка "ИНН" в 1С
Public Const A1C_INVOICE_COL = 15  ' колонка "Список счетов с оплатой" в 1С
Public Const A1C_GOOD_COL = 16  ' колонка "Товар последнего счета" в 1С


Public Const ACC1C_STAMP = "Название фирмы" ' Название фирмы используется как штамп

'~~~~~~~~~~~~~~~~~~~~~~~~ Склад ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Const STOCK_SHEET = "Склад"  ' Лист проводок по Складу
Public Const STOCK_STAMP = "Наименование продукта"
Public Const STOCK_MYCOLS = 6       ' количество моих колонок слева

Public Const STOCK_NEWSN_COL = 1    ' колонка "SN"
Public Const STOCK_CONTRADSK_COL = 2 ' колонка "ContrADSK"
Public Const STOCK_SF_SN_COL = 3    ' колонка "SF_SN"
Public Const STOCK_GOOD_COL = 4     ' колонка "Товар"
Public Const STOCK_INVOICE_COL = 5  ' колонка "Счет"
Public Const STOCK_ACC_COL = 6      ' колонка "Acc"
Public Const STOCK_CLIENT_COL = 7   ' колонка "Клиент"
Public Const STOCK_PRODUCT_COL = 8  ' колонка "Наименование продукта"
Public Const STOCK_SN_COL = 10      ' колонка "Серийный №"
Public Const STOCK_DATE_COL = 11    ' колонка "Дата"
Public Const STOCK_DELIVERY_COL = 12 ' колонка "Доставка"


'############## Заказы у Дистрибуторов или Поставщиков #################
Public Const OrderList = "Заказы"   ' Лист Заказов
Public Const OrderListStamp = "Заказ"
Public Const OrderList_MYCOLS = 5   ' количество моих колонок слева
Public Const OL_MIN_RESLINES = 30   ' Min кол-во строк пятки OL

Public Const OL_PAIDDAT_COL = 1     ' колонка "Дата оплаты"
Public Const OL_INV1C_COL = 2       ' колонка "Счет 1С"
Public Const OL_SALE1C_COL = 4      ' колонка "Продавец 1С"
Public Const OL_ACC1C_COL = 5       ' колонка "Acc1С"
Public Const OL_ORDERN_COL = 10     ' колонка "ЗаказCSD"
Public Const OL_INV_1C_COL = 11     ' "Номер счета 1С" - Фролов
Public Const OL_CSDINVN_COL = 15    ' колонка "№ счета CSD"
Public Const OL_CSDINVDAT_COL = 16  ' колонка "Дата счета CSD"

'#.#.#.#.#.#.#. Лист новых Заказов для DL .#.#.#.#.#.#.#.#.#
Public Const NewOrderList = "NewOrderList"   ' Лист Новых Заказов
Public EOL_NewOrderList     ' последняя значащая строка листа Новых Заказов

'~#~#~#~#~#~#~#~#~ Письма БТО - Заказы CSD - Склад ~#~#~#~#~#~#~#~#~
Public Const BTO_SHEET = "BTOlog"   ' Сводка по письмам БТО из CSD

Public Const BTO_DELIVERY_COL = 1   ' колонка "Доставка со склада"
Public Const BTO_PAID_DATE_COL = 2  ' колонка "Дата оплаты 1С"
Public Const BTO_INV_1C_COL = 3     ' колонка "Счет 1С"
Public Const BTO_DATE_COL = 4       ' колонка "Дата письма"
Public Const BTO_CSDATR_COL = 5     ' колонка "Заказ/Счет CSD"
Public Const BTO_SALE_COL = 6       ' колонка "Продавец"
Public Const BTO_CLIENT_COL = 7     ' колонка "Заказчик"
Public Const BTO_SN_COL = 8         ' колонка "SN"
Public Const BTO_GOOD_COL = 9       ' колонка "Товар ADSK"
Public Const BTO_MAIL_COL = 10      ' колонка "Письмо"

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
Public Const ADACC_TEL_COL = 14         ' колонка "Telephone"
Public Const ADACC_FAX_COL = 15         ' колонка "fax"
Public Const ADACC_FACTSTREET_COL = 17  ' колонка "FactStreet"
Public Const ADACC_FACTCITY_COL = 18    ' колонка "FactCity"
Public Const ADACC_FACTINDEX_COL = 19   ' колонка "FactIndex"
Public Const ADACC_FACTSTATE_COL = 20   ' колонка "FactState"
Public Const ADACC_FACTCOUNTRY_COL = 21 ' колонка "FactCountry"
Public Const ADACC_URL_COL = 22         ' колонка "web-address"
Public Const ADACC_EMAIL_COL = 23       ' колонка "email"
Public Const ADACC_CONTACT1C_COL = 24   ' колонка "Contact1C" - не загружать в DL!

'........ Лист новых связей Организаций существующих в SF и 1С ...........
Public Const AccntUpd = "AccntUpd"      ' лист новых связей Организаций SF и 1С

Public Const ACCUPD_SFID_COL = 1        ' колонка "SFaccId"
Public Const ACCUPD_1CNAME_COL = 2      ' колонка "Acc1C"
Public Const ACCUPD_INN_COL = 3         ' колонка "INN"
Public Const ACCUPD_TEL_COL = 4         ' колонка "Telephone"
Public Const ACCUPD_FAX_COL = 5         ' колонка "fax"
Public Const ACCUPD_STREET_COL = 6      ' колонка "AccStreet"
Public Const ACCUPD_CITY_COL = 7        ' колонка "AccCity"
Public Const ACCUPD_INDEX_COL = 8       ' колонка "AccIndex"
Public Const ACCUPD_STATE_COL = 9       ' колонка "AccState"
Public Const ACCUPD_COUNTRY_COL = 10    ' колонка "AccCountry"
Public Const ACCUPD_DELSTREET_COL = 11  ' колонка "factStreet"
Public Const ACCUPD_DELCITY_COL = 12    ' колонка "factCity"
Public Const ACCUPD_DELINDEX_COL = 13   ' колонка "factIndex"
Public Const ACCUPD_DELSTATE_COL = 14   ' колонка "factState"
Public Const ACCUPD_DELCOUNTRY_COL = 15 ' колонка "factCountry"
Public Const ACCUPD_URL_COL = 16        ' колонка "web-address"
Public Const ACCUPD_EMAIL_COL = 17      ' колонка "email"

'................ Лист новых Договоров C_Contr ..........................
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

'................ Лист новых Проектов O_NewOpp ..........................
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

'****************** Лист We - Словари, таблицы *************************
Public Const We = "We"  ' Лист: наши юр.лица, Продавцы, коды валют, типы товаров

Public Const WE_GOODS_COL = 1           ' Продукты (Goods)
Public Const WE_GOODS_LINCOL = 3        ' Направление деятельности
Public Const WE_GOODS_KINDCOL = 4       ' Тема & вид деятельности
Public Const WE_GOODS_ISSBSCOL = 8      ' Товар - Есть подписка
Public Const WE_GOODS_NOSBSCOL = 9      ' Товар - Нет подписки

Public Const WE_GOODS_ADSK = "Autodesk" ' Товар - Autodesk

'=============== названия листов MatchSF-1C.xlsx ========================
Public Const SF = "SF"              ' лист отчета по Платежам
Public Const SFD = "SFD"            ' лист отчета по Договорам
Public Const SFacc = "SFacc"        ' лист - список Организаций
Public Const SFopp = "SFopp"        ' лист отчета по Проектам
Public Const SForders = "SForders"  ' лист отчета по Заказам
Public Const ADSKfrSF = "ADSKfrSF"  ' лист отчета по Autodesk
    
Public Const A_Dic = "A_Dic"                ' лист - Словарь Организаций
Public Const P_Paid = "P_Paid"              ' лист новых Платежей
Public Const O_NewOpp = "O_NewOpp"          ' лист новых Проектов
Public Const C_Contr = "C_Contr"            ' лист новых Договоров
Public Const C_ContrLnk = "C_ContrLnk"      ' лист связок Договоров с Проектами
Public Const P_PaymentUpd = "P_PaymentUpd"  ' лист связей Платежей и Договоров
Public Const P_ADSKlink = "P_ADSKlink"      ' лист связок Платежей с Контрактом ADSK

' имена отчетов SF - используются как штампы
Public Const SFpayRepName = "Платежи: Сверка SF с 1С"
Public Const SFcontrRepName = "Match: SFD"          ' Договоры - Contracts
Public Const SFaccRepName = "SFacc"                 ' Организации - Accounts
Public Const SFoppRepName = "Match SFopp"           ' Проекты - Opportunities
Public Const SFadskRepName = "Match ADSK from SF"   ' Autodesk - ADSK

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

Public Const DATE_BULKY = "1.1.20"   ' дата окончания для Bulky Проектов
Public Fruitful         ' счетчик полезных результатов

Public ExRespond As Boolean ' если False - завершение работы

'########################## лист лидов MS CSIT ###########################
Public Const CSIT_MS = "CSIT_MS"            ' лист Лидов CSIT по Microsoft
Public EOL_CSIT_MS                          ' последняя строка листа CSIT_MS

Public Const CSIT_MS_IDSF_COL = 1           ' Колонка "IdSF" - Id соответсветвствующих
                                            '                    Организаций в SF
Public Const CSIT_MS_NAME_COL = 2           ' Колонка "Название" - имя организации
Public Const CSIT_MS_ADDR_COL = 5           ' Колонка "Адрес"
Public Const CSIT_MS_TEL_COL = 8            ' Колонка "телефон"
Public Const CSIT_MS_FAX_COL = 9            ' Колонка "Факс"
Public Const CSIT_MS_URL_COL = 10           ' Колонка "Web-site"
Public Const CSIT_MS_EMAIL_COL = 11         ' Колонка "E-mail"
Public Const CSIT_MS_INN_COL = 19           ' Колонка "ИНН"
Public Const CSIT_MS_STAMP = "Первое Независимое Рейтинговое Агентство"
Public Const CSIT_MS_NAME = "Название"      ' заголовок колонки

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
    SN      As String           ' Серийный номер
    Contr   As String           ' Контракт Autodesk / Agreement
    AccN    As String           ' Customer Account Number, CSN на PartnerCenter
    accName As String           ' Имя клиента на PartnerCenter
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


