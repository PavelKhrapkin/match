Attribute VB_Name = "Declarations"
'-------------------------------------------------------------------
' Declarations - ���������� ��������, ������������ � match 2.1
'
'  12.9.13

Option Explicit

'---- ������� "��������" ��� "Downloads" ----
Public Const DownloadDir = "C:\Users\Pavel_Khrapkin\Downloads\"
'============================ ����� DBs ===============================
Public DirDBs As String    ' ������� ������ DBs �� 'match.xlsm'!We
' ��������������� ���� 'match_environment.xlsx' �������� DirDBs
Public Const F_match_environment = "C:\match_environment.xlsx"

Public Const F_MATCH = "match.xlsm"
Public Const F_1C = "1C.xlsx"
Public Const F_SFDC = "SFDC.xlsx"
Public Const F_ADSK = "ADSK.xlsm"
Public Const F_STOCK = "Stock.xlsx"
Public Const F_TMP = "W_TMP.xlsm"   '.xlsm ������, ��� �������� �������

Public DB_MATCH As Workbook '������ � ������� match
Public DB_1C As Workbook    '������ 1C
Public DB_SFDC As Workbook  '������ Salesforce
Public DB_ADSK As Workbook  '������ Autodesk
Public DB_STOCK As Workbook '������ �� ������ � �������
Public DB_TMP As Workbook   '������� ���� ��� ������ "�������" ������ � WP

'==================== �������� �� ����� Process ====================
Public Const Process = "Process"    '������� � �������� ���� ���������
Public Const PROC_START = "<*>ProcStart"    ' ������ ��������
Public Const PROC_END = "<*>ProcEnd"        ' ����� ��������

Public Const PROC_TIME_COL = 1      '������� 1 - ���� � ����� ���������� ����
Public Const PROC_STEP_COL = 2      '������� Step - ��� ��������
Public Const PROC_PREVSTEP_COL = 3  '������� PrevStep
Public Const PROC_STEPFILE_COL = 4  '������� "� �����" - ����, ��� ��� ����
Public Const PROC_STEPDONE_COL = 5  '������� Done - ��� ��������
Public Const PROC_PAR1_COL = 6      '������� �������� 1 - ��� ��������
Public Const PROC_PAR2_COL = PROC_PAR1_COL + 1
Public Const PROC_NAME_COL = PROC_PAR1_COL ' ��� �������� � Par1
Public Const PROC_REP1_COL = 11     '�������  - ��� ����� - �������� 1

Public Const PROCESS_NAME_COL = 7  '������� - ��� �������� ��������
Public Const STEP_NAME_COL = 9     '������� - ��� �������� ����
Public PublicProcName As String
Public PublicStepName As String

'-- �������� � ������������� ���� 1C.xlsm
Public Const REP_1C_P_LOAD = "�������� �������� �� 1�"
Public Const REP_1C_P_PAINT = "��������� �������� 1�"
Public Const REP_1C_SFACCFIL = "���������� ������� 1 ��� �������� 1�"

Public Const REP_1C_�_LOAD = "�������� ��������� �� 1�"
Public Const REP_1C_�_PAINT = "��������� ��������� 1�"
Public Const REP_1C_A_LOAD = "�������� ����������� �������� �� 1�"

'-- �������� � ������������� ������� �� SalesForce � ���� SDFC.xlsm
Public Const REP_SF_LOAD = "�������� �������� �� Salesforce - SF"
Public Const REP_SFD_LOAD = "�������� ��������� �� Salesforce - SFD"
    
'-- ���������� ����������, ����� ��� ���� �������
Public RepName As String    ' ��� ��������������� ������
Public Doing As String      ' ������ � Application.StatusBar - ��� ������ ������
Public AllCol As Integer    ' ���������� ������� � ������� ������
Public Lines As Integer     ' ���������� ����� ��������/������ ������
Public LinesOld As Integer  ' ���������� ����� ������� ������

Public Fruitful As Integer  ' ������� �������� �����������

Public ExRespond As Boolean ' ���� False - ���������� ������
Public Silent As Boolean    ' ���� True - �� �������� ��������� �� �������

'=============== ���� DB_MATCH - ���� match.xlsm ==============
Public Const Header = "Header"              ' ���� ���������� (�����) � ���������
Public Const A_Dic = "A_Dic"                ' ���� - ������� �����������

Public EOL_PaySheet     ' ��������� ������ ������ 1� �� �������� ��� �����
Public EOL_DogSheet     ' ��������� ������ ������ 1� �� ��������� ��� �����
Public EOL_SF           ' ��������� ������ ������ �� �������� SF ��� �����
Public EOL_SFD          ' ��������� ������ ������ �� ��������� SFD ��� �����
Public EOL_SFopp        ' ��������� ������ ������ �� �������� � SFopp ��� �����
Public EOL_SFacc        ' ��������� ������ SFacc �� ������������ ��� �����
Public EOL_SForders     ' ��������� ������ ������ �� ������� SForders ��� �����
Public EOL_ADSKfrSF     ' ��������� ������ ������ Autodesk �� SF
Public EOL_Acc1C        ' ��������� ������ ����������� ����������� 1�
Public EOL_NewOpp       ' ��������� ������ ����� ����� ��������
Public EOL_NewPay       ' ��������� ������ ����� ����� ��������
Public EOL_NewContr     ' ��������� ������ ����� ����� ���������
Public EOL_NewSN        ' ��������� ������ ����� ����� SN Autodesk
Public EOL_ContrLnk     ' ��������� ������ ����� ����� ��� ����������� ���������
Public EOL_PaymentUpd   ' EOL ����� ��������� �������� - ������ ������� � ���������
Public EOL_DIC          ' ��������� ������ ����� ������� �����������
Public EOL_AdAcc        ' ��������� ������ ����� ����� ����������� ��� �������� � SF
Public EOL_AccntUpd     ' ��������� ������ ����� ����� ������ ����������� Sf � 1� ��� �������� � SF
Public EOL_ADSKlnkPay   ' ��������� ������ ����� ����� ������ �������-������� ADSK
Public EOL_SFlnkADSK    ' ��������� ������ ����� ������ �������-������� ADSK
Public EOL_ADSK         ' ��������� ������ ����� ������ �� ADSK.xlsx
Public EOL_Stock        ' ��������� �������� ������ ����� �� ������
Public EOL_BTO          ' ��������� �������� ������ ����� BTOlog
Public EOL_OrderList    ' ��������� �������� ������ ����� �������

Public Const DATE_BULKY = "1.1.20"  ' ���� ��������� ��� Bulky ��������
Public Const BIG = 77777            ' ������� ����� ��� ������ ������

'------------- match TOC - ���������� ������� � ���� ������ ----------
Public Const TOC = "TOCmatch"           ' ���������� ������ ���� ������ - ��� ������
Public Const TOC_F_DIR_COL = 10         ' � TOCmatch ������� Path ������ DBs - DirDBs

Public Const TOC_DATE_COL = 1           ' ���� � ����� �������� ������
Public Const TOC_REPNAME_COL = 2        ' ��� ������ � ���� ������
Public Const TOC_EOL_COL = 3            ' EOL ������ ��� �����
Public Const TOC_MYCOL_COL = 4          ' MyCol - ����� ���.������� �����
Public Const TOC_RESLINES_COL = 5       ' ����� ����� ����� ResLines
Public Const TOC_MADE_COL = 6           ' Made - ����������� ��� �� �����
Public Const TOC_MAXDAYS_COL = 7        ' ������������� - max ���� � ���������� ������
Public Const TOC_REPFILE_COL = 8        ' ��� �����, ����������� �����
Public Const TOC_SHEETN_COL = 9         ' ��� �����, ����������� ����� (��������)
Public Const TOC_STAMP_COL = 10         ' �����
Public Const TOC_STAMP_TYPE_COL = 11    ' ��� ������: ������ (=) ��� ��������� (I)
Public Const TOC_STAMP_R_COL = 12       ' ������ ������: (��������, +EOL)
Public Const TOC_STAMP_C_COL = 13       ' ������� ������: (��������, +MyCol)
Public Const TOC_CREATED_COL = 14       ' ���� � ����� �������� ������
Public Const TOC_PAR_1_COL = 15         ' ������� ������ �������� 1
Public Const TOC_PAR_2_COL = 16         ' ������� ������ �������� 2
Public Const TOC_PAR_3_COL = 17         ' ������� ������ �������� 3
Public Const TOC_PAR_4_COL = 18         ' ������� ������ �������� 4
Public Const TOC_PAR_5_COL = 19         ' ������� ������ �������� 5
Public Const TOC_PAR_6_COL = 20         ' ������� ������ �������� 6
                        '--- ��� ��������� ���������� ������, � �� ��� ������
Public Const TOC_FRDATEROW_COL = 22     ' ������ ������ ��� FrDate
Public Const TOC_TODATEROW_COL = 23     ' ������ ������ ��� ToDate
Public Const TOC_DATECOL_COL = 24       ' ������� ������ ��� ����
Public Const TOC_NEW_FRDATE_COL = 25    ' FrDate ������ ������
Public Const TOC_NEW_TODATE_COL = 26    ' ToDate ������ ������

Public Const TOC_PARCHECK_COL = TOC_PAR_1_COL   ' ������ ���.������
Public Const TOC_INSHEETN = TOC_PAR_2_COL       ' ������� - InSheetN
Public Const TOC_FORMNAME = TOC_PAR_2_COL       ' ������� - ��� �������
Public Const TOC_FORMSUMMARY = TOC_PAR_3_COL    ' ������� - ��� ������� �����
Public Const TOC_FRDATE_COL = TOC_PAR_4_COL     ' �������- FrDate ������� ������
Public Const TOC_TODATE_COL = TOC_PAR_5_COL     ' �������- ToDate ������� ������
Public Const TOC_REPLOADER_COL = TOC_PAR_6_COL  ' �������- Loader ������

Public Const TOCstart = 4       ' ������ ������ TOCmatch - ������� ��������� �����
Public Const TOCrepLines = 8    ' ������ ����� ����������. ����� ��������� �����
    
Type TOCmatch
    iTOC As Long        '=����� ������ � ������� TOCmatch - Read ONLY!
    Dat As Date         '=Now   - ���� � ����� �������� ������
    Name As String      '�� ���.- ��� ������ � ���� ������
    EOL  As Long        '=���.������ MoveToMatch - EOL ������ ��� �����
    MyCol As Long       '=���.InsMyCol   - MyCol - ����� ���.������� �����
    ResLines As Long    '=���.InsSmmary  - ����� ����� � ����� ������ ����� EOL
    Made As String      '=���.������ ��� - Made    - ����������� ��� �� �����
    RepFile As String   '�� ���.-  ���� DB � �������
    SheetN As String    '�� ���.-  ��� �����, ����������� �����
''''    Stamp As String     '�� ���.-  ������ - �����
''''    StampType As String '�� ���.-  ��� ������: ������ (=) ��� ��������� (I)
''''    StampR As Long      '�� ���.-  ������ ������: (+EOL)
''''    StampC As Long      '�� ���.-  ������� ������: (+MyCol)
    CreateDat As Date   '=���.������ MoveToMatch � NewSheet -  ���� � ����� �������� ������
''''    ParChech As String  '�� ���.-   <>"" - ����.������ - ���.�����
''''    Loader As String    '�� ���.-   Loader ������
    FormName As String  '�� ���.-   �������� ����� - �������
End Type

Public RepTOC As TOCmatch   ' ��������� TOCmatch
''''''''''''Public RepMatch As TOCmatch ' ��������� TOCmatch ��� ������ � match.xlsm
''''''''''''Public Rep1C As TOCmatch    ' ��������� TOCmatch ��� ������ � 1�.xlsm
''''''''''''Public RepSF As TOCmatch    ' ��������� TOCmatch ��� ������ � SFDC.xlsm
''''''''''''Public RepADSK As TOCmatch  ' ��������� TOCmatch ��� ������ � ADSK.xlsm
''''''''''''Public RepStock As TOCmatch ' ��������� TOCmatch ��� ������ � Stock.xlsm

Public Const REP_LOADED = "Loaded"            ' MoveToMatch: ����� �������� � ���� DB

'=============== ���� DB_SFDC - ���� SFDC.xlsm ==============
Public Const SF = "SF"              ' ���� ������ �� ��������
Public Const SFD = "SFD"            ' ���� ������ �� ���������
Public Const SFacc = "SFacc"        ' ���� - ������ �����������
Public Const SFcont = "SFcont"      ' ���� ������ �� ���������
Public Const SFopp = "SFopp"        ' ���� ������ �� ��������
Public Const SForders = "SForders"  ' ���� ������ �� �������
Public Const ADSKfrSF = "ADSKfrSF"  ' ���� ������ �� Autodesk

Public Const SFresLines = 6 '������ ����� ������� SalesForce

'------------- SF - ����� Saleforce �� �������� --------------------
Public Const SF_MATCH_COL = 1      ' �������-Match - ����� ������ ������� 1�
Public Const SF_COD_COL = 2        ' �������-��� ������� � SF
Public Const SF_ACC1C_COL = 3      ' �������-"��� ����������� � 1�"
Public Const SF_INV_COL = 8        ' �������-"����� �����"
Public Const SF_SAIL_COL = 9       ' �������- ���������, ������������� �� ������
Public Const SF_PAYID_COL = 18     ' �������- Id ������� � SF
    
'------------- SFD - ����� Saleforce �� ��������� ------------------
Public Const SFD_COD_COL = 2       ' �������-��� "���/���" � SFD
Public Const SDF_CONTRN_COL = 3    ' ������� ����� �������� � SF
Public Const SFD_ACC1C_COL = 6     ' ������� "��� ����������� � 1�"
Public Const SFD_DATESTART_COL = 8 ' ������� "���� ������ ��������"
Public Const SFD_DATEEND_COL = 9   ' ������� "���� ��������� ��������"
Public Const SFD_SCAN_COL = 10     ' ������� "�������������"
Public Const SFD_OWNER_COL = 12    ' ������� "���."
Public Const SFD_STATUS_COL = 13   ' ������� "������"
Public Const SFD_CONTRID_COL = 15  ' ������� "��� ��������"
Public Const SFD_OPPN_COL = 17     ' ������� OppN � SFD
Public Const SFD_OPPID_COL = 18    ' ������� OppId
Public Const SFD_VALUE_COL = 20    ' ������� ����� ��������

'--------------- SFopp ������� ----------------------------------------
Public Const SFOPP_OPPID_COL = 1        ' ������� "��� �������"
Public Const SFOPP_OPPN_COL = 2         ' ������� "����� �������"
Public Const SFOPP_OPPNAME_COL = 3      ' ������� "��� �������"
Public Const SFOPP_ACC1C_COL = 4        ' ������� ����������� 1� ����� SFopp
Public Const SFOPP_PROBABILITY_COL = 5  ' ������� "����������� (%)"
Public Const SFOPP_CRDAT_COL = 8        ' ������� ���� �������� ������� SFopp
Public Const SFOPP_CLOSEDATE_COL = 9    ' ������� "���� ��������" �������
Public Const SFOPP_SALE_COL = 10        ' ������� "������������� �� ������"
Public Const SFOPP_TO_PAY_CUR_COL = 14  ' ������� "��� �� �������� ������" ����� SFopp
Public Const SFOPP_TO_PAY_VAL_COL = 15  ' ������� "��� �� ��������" ����� SFopp
Public Const SFOPP_LINE_COL = 17        ' ������� "����������� ������������"
Public Const SFOPP_TYP_COL = 18         ' ������� "����&��� ������������" �������

'--------------- SFacc ����������� ------------------------------------
Public Const SFACC_ACCNAME_COL = 1      ' ������� "��� o����������" � SF
Public Const SFACC_ACC1C_COL = 2        ' ������� "��� ����������� � 1�"
Public Const SFACC_IDACC_COL = 3        ' ������� "��� �����������"
Public Const SFACC_STREET_COL = 4       ' ������� "����� ��� �����"
Public Const SFACC_CITY_COL = 5         ' ������� "����� ��� �����"
Public Const SFACC_STATE_COL = 6        ' ������� "������� ��� �����"
Public Const SFACC_INDEX_COL = 7        ' ������� "������ ��� �����"
Public Const SFACC_COUNTRY_COL = 8      ' ������� "������ ��� �����"

'--------------- SFADSK - ����� �� Autodesk �� SF ------------------------
Public Const SFADSK_CONTRACT_COL = 19   ' ������� "����� ��������� Autodesk"
Public Const SFADSK_CONTRID_COL = 2     ' ������� "Ko� ��������� Autodesk" = Id SF
Public Const SFADSK_CONTRSTAT_COL = 3   ' ������� "Status" - ������ ��������� Autodesk
Public Const SFADSK_SN_COL = 4          ' ������� "SN �������� Autodesk"
Public Const SFADSK_DESCRIPTION_COL = 5 ' ������� "��� �������� Autodesk"
Public Const SFADSK_SEATS_COL = 6       ' ������� "Seats" - ����
Public Const SFADSK_STATUS_SN_COL = 7   ' ������� "������ SN"
Public Const SFADSK_ACCN_COL = 8        ' ������� "Account_#"
Public Const SFADSK_ACCNAME_COL = 9     ' ������� "Account_Name ADSK"
Public Const SFADSK_CONTRSTARTDAT_COL = 11 ' ������� "Contract Start Date"
Public Const SFADSK_CONTRENDDAT_COL = 12   ' ������� "Contract End Date"
Public Const SFADSK_SALE_COL = 15       ' ������� "��������"
Public Const SFADSK_ISSBS_COL = 20        ' ������� "���� ��������"
Public Const SFADSK_SNREGDAT_COL = 21   ' ������� "���� ���������"
Public Const SFADSK_ACC1C_COL = 23      ' ������� "��� o���������� � 1�"
Public Const SFADSK_CM_NAME_COL = 24    ' ������� "������� ��: ������ ���"
Public Const SFADSK_CM_MAIL_COL = 25    ' ������� "CM ����"
Public Const SFADSK_CM_TEL_COL = 26     ' ������� "CM �������"
Public Const SFADSK_DEPLOYMENT_COL = 27 ' ������� "Deployment"

Public Const SFADSK_SN_REGISTERED = "Registered"   ' "������ SN" Registered

'- - - - - - - - SF_PA = "SF_P_ADSK_LINK" - - - - - - - - - - - - - -
Public Const SF_PA = "SF_PA"        ' ���� ������ ������-�������� ADSK

Public Const SFPA_PAYID_COL = 2     ' ������� "������: ��� ������"
Public Const SFPA_ADSKID_COL = 3    ' ������� "��� ��������� ADSK"

'. . . .  ��������� ����� ���� SNatr  - ����, ����������� �� ADSK.xlsx . . . .
'----- ���������� ���� ADSK.xlsx --------------------
Public Const TOC_ADSK = "TOC_ADSK"
Public Const TOCADSK_REPNAME_COL = 3      '���� - ���/��� ������
Public Const TOCADSK_REPRANGE_COL = 5     '���� - ���� (Range)

Public ADSKrep                      ' ��� ������ �� ADSK.xlsx
Public Const ADSK_HdrMapSize = 20           ' ������ ADSK_RepMap
Public ADSK_RepMap(1 To ADSK_HdrMapSize)    ' ��������� ����� ������ ADSKrep
Public ADSK_HDR_Map(1 To ADSK_HdrMapSize)   ' ��������� ����� ������ �� ADSK

Public SN_SF As SNatr               ' SN �� ����� ���������� �� SF
Public SN_ADSK As SNatr             ' SN �� ����� ���������� - ������ �� ADSK.xlsx

Public Const ADSK_CONTR_COL = 1     ' ������� "ContrADSK" - �������� Autodesk / Agreement
Public Const ADSK_ACCN_COL = 2      ' ������� "AccN" - CSN �� PartnerCenter
Public Const ADSK_ACCNAME_COL = 3   ' ������� "AccName" - ��� ������� �� PartnerCenter
Public Const ADSK_C_START_COL = 4   ' ������� "ContrStartDate" - ���� ������ ��������� ADSK
Public Const ADSK_C_END_COL = 5     ' ������� "ContractEndDate" - ���� �������� ���������             ' ���� ��������� ��������� ADSK
Public Const ADSK_C_STAT_COL = 6    ' ������� "Contr_Status" - ������ ��������� ADSK
Public Const ADSK_CM_NAME_COL = 8   ' ������� ��� �� - �������� �� SN
Public Const ADSK_CM_L_NAME_COL = 9 ' ������� ��� �� - �������
Public Const ADSK_CM_F_NAME_COL = 10 ' ������� ��� �� - ���
Public Const ADSK_CM_MAIL_COL = 11  ' ������� email CM
Public Const ADSK_CM_TEL_COL = 12   ' ������� ������� ��
Public Const ADSK_SN_COL = 13       ' ������� "SN"
Public Const ADSK_SN_DESCR_COL = 14 ' ������� �������� �������� Autodesk
Public Const ADSK_SN_STATUS_COL = 15 ' ������� ������ SN: Registered, Upgrade...
Public Const ADSK_DEPL_COL = 16     ' ������� "Deployment" �������� ������� ��� ���������
Public Const ADSK_SEATS_COL = 17    ' ������� ����� ��������/���������� ����
Public Const ADSK_SBS_COL = 18      ' ������� "Subscriptions" = Sbs - ���� �� ��������              ' TRUE ���� �������� �������
Public Const ADSK_YEAR_COL = 19     ' ������� "Release_Year" - ��� ������ �������� ADSK
Public Const ADSK_REGDAT_COL = 20   ' ������� ���� ����������� SN

'. . . .  ADSK_C - ���� ���������� ADSK �� PartnerCenter . . . .
Public Const ADSK_C = "ADSK_C"      ' ���� ���������� ADSK

Public Const ADSK_C_STAMP = "Contract #"

Public Const ADSK_C_CONTR_COL = 2   ' ������� "Contract #"
Public Const ADSK_C_ENDDATE_COL = 3 ' ������� "Contract End Date"
Public Const ADSK_C_ACCN_COL = 11   ' ������� "End Customer CSN"

'-------------- ������� - ����� �� 1� ---------------------------------
Public Const PAY_SHEET = "�������"  ' ���� ��������

''' Public Const PAY_RESLINES = 3       ' ���-�� ����� ����� ������ �� ��������
''' Public Const PAY_MYCOLS = 5         ' ���������� ���� ������� �����

Public Const PAYISACC_COL = 1       ' ������� =1, ���� ����������� ���� � SF
Public Const PAYINSF_COL = 2        ' ������� =1, ���� ������ ���� � SF
Public Const PAYIDSF_COL = 3        ' ������� - Id ������� � SF
Public Const PAYADSK_COL = 4        ' ������� - ����� ������� � ADSK � SF
Public Const PAYCODE_COL = 5        ' ������� ���� �������
'----
Public Const PAYDOC_COL = 6         ' ������� "����.���."
Public Const PAYDATE_COL = 7        ' ������� "���� ����.�����"
Public Const PAYINVOICE_COL = 8     ' ������� "����"
Public Const PAYACC_COL = 9         ' ������� "������"
Public Const PAYINVVAL_COL = 10     ' ������� "����� ����� � ���"
Public Const PAYRUB_COL = 18        ' ������� "���� ���"
Public Const PAYGOOD_COL = 19       ' ������� �����
Public Const PAYFIRM_COL = 20       ' ������� �����
Public Const PAYSALE_COL = 22       ' ������� ��������
Public Const PAYDOGOVOR_COL = 25    ' ������� �������
Public Const PAYOSNDOGOVOR_COL = 26 ' ������� ���.�������
Public Const PAYDOGVAL_COL = 27     ' ������� ����� ��������
Public Const PAYDOGCUR_COL = 28     ' ������� ������ ��������
Public Const PAYINV_COL = 30        ' ������� CX-XXX - ���� ��� ����
Public Const PAYGOODTYPE_COL = 31   ' ������� ��� ������

'''Public Const Stamp1Cpay1 = "����. ���.", Stamp1Cpay2 = "���� ����. �����"
'-------------- �������� - ����� �� 1� ---------------------------------
Public Const DOG_SHEET = "��������" ' ���� ��������� ��� ������ ConctAnalitics

Public Const DOGRES = 7             ' ���-�� ����� ����� ������ �� ���������

Public Const DOGISACC_COL = 1      ' ������� =1, ���� ����������� ���� � SF
Public Const DOGSFSTAT_COL = 2     ' ������� ������� �������� � SF
Public Const DOGIDSF_COL = 3       ' ������� Id �������� � SF
Public Const DOGPAID1C_COL = 5     ' ������� ��������1C "��������"
Public Const DOGISINV1C_COL = 6    ' ������� ��������1C "�� ������"
Public Const DOG1CSCAN_COL = 7     ' ������� "�������������"
Public Const DOGCOD_COL = 8        ' ������� ���� �������� ���� "���/������"
'----
Public Const DOG1C_COL = 10        ' ������� "����� ��������" � ������ 1�
Public Const DOG1CDAT_COL = 12     ' ������� ���� ����������
Public Const DOG1CFIRM_COL = 13    ' ������� �����
Public Const DOG1CSALE_COL = 14    ' ������� ��������
Public Const DOG1CACC_COL = 15     ' ������� ��������
Public Const VENDOR1C_COL = 16     ' ������� ��������� ����� ��������� 1�
Public Const DOGTYPE_COL = 17      ' ������� ��� ��������
Public Const DOG1C_MAINDOG_COL = 18 ' ������� "�������� �������"
Public Const DOGSUM1C_COL = 20     ' ������� ����� (����)
Public Const DOGINVC_COL = 21      ' ������� ����� �� ������
Public Const DOGPAID_COL = 22      ' ������� ����� �� ��������
Public Const DOGCUR1C_COL = 23     ' ������� ������ ��������

Public Const DOG_STAT_DRAFT = "��������"
Public Const DOG_STAT_OPEN = "������"
Public Const DOG_STAT_CLOSED = "������"
Public Const DOG_STAT_CANCEL = "�� ���������"
Public Const DOG_STAT_NOTINSF = "��� � SF"

'-------------- ���������� ����������� - ����� �� 1� -------------------
Public Const Acc1C = "������ �������� 1C"   ' ���������� �������� 1�

Public Const ACC1C_RES = 3      ' ����� ����������� ����������� 1�

Public Const A1C_ISACC_COL = 1  ' ������� =1, ���� ����������� ���� � SF
'---
Public Const A1C_NAME_COL = 5   ' ������� "�������� �����" � 1�
Public Const A1C_CON_COL = 6    ' ������� "�������" � 1�
Public Const A1C_TEL_COL = 7    ' ������� "�������"
Public Const A1C_ADR_COL = 8    ' ������� "�����" � 1�
Public Const A1C_FACTADR_COL = 9 ' ������� "����. �����" � 1�
Public Const A1C_INN_COL = 10   ' ������� "���" � 1�
Public Const A1C_INVOICE_COL = 17  ' ������� "������ ������ � �������" � 1�
Public Const A1C_GOOD_COL = 18  ' ������� "����� ���������� �����" � 1�

'~~~~~~~~~~~~~~~~~~~~~~~~ ����� ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Const STOCK_SHEET = "�����"  ' ���� �������� �� ������
Public Const STOCK_STAMP = "������������ ��������"
Public Const STOCK_MYCOLS = 6       ' ���������� ���� ������� �����

Public Const STOCK_NEWSN_COL = 1    ' ������� "SN"
Public Const STOCK_CONTRADSK_COL = 2 ' ������� "ContrADSK"
Public Const STOCK_SF_SN_COL = 3    ' ������� "SF_SN"
Public Const STOCK_GOOD_COL = 4     ' ������� "�����"
Public Const STOCK_IDSF_COL = 5     ' ������� "IdSF"
Public Const STOCK_IDSFORDER_COL = 6 ' ������� "IdSForder"
Public Const STOCK_INVOICE_COL = 7  ' ������� "����"
Public Const STOCK_ACC_COL = 8      ' ������� "Acc"
Public Const STOCK_CLIENT_COL = 9   ' ������� "������"
Public Const STOCK_PRODUCT_COL = 10  ' ������� "������������ ��������"
Public Const STOCK_SN_COL = 12      ' ������� "�������� �"
Public Const STOCK_DATE_COL = 13    ' ������� "����"
Public Const STOCK_OLDSN_COL = 15   ' ������� "������ S/N"

'~ ~ ~ ~ ~ ~ ~ ~ ~ ������ ��� �� ����������� �� ������ ~ ~ ~ ~ ~ ~ ~ ~ ~
Public Const BTOmails = "BTOmails"  ' �������� ������ ��� �� ����� BTOmails.txt
Public Const BTO_SHEET = "BTOlog"   ' ������ ��� ��� �������� �� ������

Public Const BTO_DATE_COL = 1       ' ������� "���� ������"
Public Const BTO_INVCSD_COL = 2     ' ������� "�����/���� CSD"
'''Public Const BTO_INV1C_COL = 3      ' ������� "���� 1�" � ����� �����
'''Public Const BTO_SALE_COL = 4       ' ������� "��������"
'''Public Const BTO_CLIENT_COL = 5     ' ������� "��������"
Public Const BTO_SN_COL = 7         ' ������� "SN"
''Public Const BTO_GOOD_COL = 7       ' ������� "����� ADSK"
''Public Const BTO_SEATS_COL = 8      ' ������� "����"
Public Const BTO_UNIQUE_COL = 11    ' ������� "UNIQUE"
''Public Const BTO_MAIL_COL = 9       ' ������� "������"

'############## ������ � ������������� ��� ����������� #################
Public Const ORDER_SHEET = "������" ' ���� �������

Public Const OL_IDSF_COL = 1        ' ������� "IdSF"
Public Const OL_GOOD_COL = 14       ' ������� "TMPgoodType" - ��������� ������ � ������
Public Const OL_INV1C_COL = 15      ' ������� "TMPinv1C" - ���� 1� � ������
Public Const OL_CUSTOMER_COL = 16   ' ������� "TMPcustomer" - ������ � ������
Public Const OL_SALES_COL = 17      ' ������� "TMPsalesRep" - �������� � ������
Public Const OL_ORDERNUM_COL = 18   ' ������� "OrderNumber" - �������� � ������
Public Const OL_CSDINVDAT_COL = 20  ' ������� "CSDinvDate" - ���� ����� CSD � ������

'................. ���� ����� ����������� A_Acc ..................
Public Const A_Acc = "A_Acc"            ' ���� ����� �����������

Public Const ADACC_NAME_COL = 1         ' ������� "AccountName"
Public Const ADACC_1CNAME_COL = 2       ' ������� "Acc1CName"
Public Const ADACC_INN_COL = 4          ' ������� "INN"
Public Const ADACC_STREET_COL = 5       ' ������� "AccStreet"
Public Const ADACC_CITY_COL = 6         ' ������� "AccCity"
Public Const ADACC_INDEX_COL = 7        ' ������� "AccIndex"
Public Const ADACC_STATE_COL = 8        ' ������� "AccState"
Public Const ADACC_COUNTRY_COL = 9      ' ������� "AccCountry"
Public Const ADACC_TEL_COL = 14         ' ������� "�������"
Public Const ADACC_FACTSTREET_COL = 17  ' ������� "FactStreet"
Public Const ADACC_FACTCITY_COL = 18    ' ������� "FactCity"
Public Const ADACC_FACTINDEX_COL = 19   ' ������� "FactIndex"
Public Const ADACC_FACTSTATE_COL = 20   ' ������� "FactState"
Public Const ADACC_FACTCOUNTRY_COL = 21 ' ������� "FactCountry"
Public Const ADACC_CONTACT1C_COL = 22   ' ������� "Contact1C" - �� ��������� � DL!

'........ ���� ����� ������ ����������� ������������ � SF � 1� ...........
Public Const AccntUpd = "AccntUpd"      ' ���� ����� ������ ����������� SF � 1�

Public Const ACCUPD_SFID_COL = 1        ' ������� "SFaccId"
Public Const ACCUPD_1CNAME_COL = 2      ' ������� "Acc1C"

'................ ���� ����� �������� NewPayment ..........................
Public Const NEW_PAYMENT = "NewPayment" ' ���� ����� ��������

Public Const NEWPAY_DOC_COL = 1         ' Doc - ��� ���������� ���������
Public Const NEWPAY_DATE_COL = 2        ' Date - ���� �������
Public Const NEWPAY_INV_COL = 3         ' Invoice - ����

'................ ���� ����� ��������� NewContract ..........................
Public Const NEW_CONTRACT = "NewPayment" ' ���� ����� ���������

Public Const NEWDOG_DOGOVOR_COL = 1     ' Dogovor - ��� ��������
Public Const NEWDOG_DATE_COL = 2        ' Date - ���� ����������
Public Const NEWDOG_WE_COL = 3          ' We - ���� �����
Public Const NEWDOG_IDWE_COL = 4        ' IdWe
Public Const NEWDOG_SALESMAN_COL = 5    ' Salesman - ��������
Public Const NEWDOG_IDOWNER_COL = 6     ' IdOwner - Id ��������
Public Const NEWDOG_CLIENT_COL = 7      ' Client - ��������
Public Const NEWDOG_IDACC_COL = 8       ' IdAcc - Id ���������
Public Const NEWDOG_TYPE_COL = 9        ' Type - ContractType - ��� ��������
Public Const NEWDOG_MAINDOG_COL = 10    ' MainContract - �������� �������
Public Const NEWDOG_CUSTCONTR_COL = 11  ' CustomerContractNumber
Public Const NEWDOG_VALUE_COL = 12      ' Value - �����(����)
Public Const NEWDOG_CURRENCY_COL = 13   ' Currency - ������
Public Const NEWDOG_CURRENCYCOD_COL = 14 ' CurrencyCod - ��� ISO ������
Public Const NEWDOG_RATE_COL = 15       ' CurrencyExchangeRate - ���� ������
Public Const NEWDOG_RATEDATE_COL = 16   ' ���� ������ �� ����
Public Const NEWDOG_ACCSTREET_COL = 17  ' �����
Public Const NEWDOG_ACCCITY_COL = 18    ' �����
Public Const NEWDOG_ACCINDEX_COL = 19   ' ������
Public Const NEWDOG_ACCSTATE_COL = 20   ' �������
Public Const NEWDOG_ACCCOUNTRY_COL = 21 ' ������

Public Const DOG_UPDATE = "ContractUpdate" ' ���� ��������� (������ � Opp) ���������
'................ ���� ����� �������� NewOpp ..........................
Public Const NEW_OPP = "NewOpp"         '"�������" ���� ����� ��������

Public Const NEWOPP_ACCID_COL = 1       ' AccId
Public Const NEWOPP_OPPNAME_COL = 2     ' OppName
Public Const NEWOPP_CLOSDATE_COL = 3    ' CloseDate
Public Const NEWOPP_OWNERID_COL = 4     ' OwnerId
Public Const NEWOPP_VALUE_COL = 5       ' Value
Public Const NEWOPP_ISOCUR_COL = 6      ' ISOcur
Public Const NEWOPP_TYPOPP_COL = 7      ' TypOpp                    Type = ���
Public Const NEWOPP_STAGENAME_COL = 8   ' StageName
Public Const NEWOPP_LINE_COL = 9        ' OpportunityLineOfActivity ����������� ������������
Public Const NEWOPP_KIND_COL = 10       ' OpportunityActivityKind   ����&��� ������������
Public Const NEWOPP_BUDDY_COL = 11      ' Buddy
Public Const NEWOPP_OPPUNIQ_COL = 12    ' OppUniq

'................ ���� ����� �������� ������� Autodesk .................
Public Const NewSN = "NewSN"            ' ���� NewSN

Public Const NEWSN_SN_COL = 1           ' SN
Public Const NEWSN_CONTRADSK_COL = 2    ' ContrADSK
Public Const NEWSN_ACCN_COL = 3         ' ����

'................ ���� ��������� �������� ������� Autodesk .................
Public iActive As Long '- ����� ������ � SNactive
Public iUpdate As Long '- ����� ������ � SNupdate
Public nActive As Long '= ������� Registered
Public nUpdate As Long '= ������� Update
Public nToClean As Long '=������� "����������� �������"

'.............. ���� SNtmp - SN ��  License Inquiry PC� Autodesk ...........
Public Const SNTMP_SN_COL = 1           ' SN
Public Const SNTMP_ACTDATE_COL = 9      ' Most Recent Activation Date
Public Const SNTMP_STATUS_COL = 12      ' ������ SN
Public Const SNTMP_STATUS_REGISTERED = "Registered"
    
'----------------------------------------------------------------------------
Public Const WP = "WP"                  ' ���� WP

'****************** ���� We - �������, ������� *************************
Public Const We = "We"  ' ����: ���� ��.����, ��������, ���� �����, ���� �������

Public Const WE_GOODS_COL = 1           ' �������� (Goods)
Public Const WE_GOODS_LINCOL = 3        ' ����������� ������������
Public Const WE_GOODS_KINDCOL = 4       ' ���� & ��� ������������
Public Const WE_GOODS_ISSBSCOL = 8      ' ����� - ���� ��������
Public Const WE_GOODS_NOSBSCOL = 9      ' ����� - ��� ��������

Public Const WE_GOODS_ADSK = "Autodesk" ' ����� - Autodesk

'########################## ���� ����� MS CSIT ###########################
Public Const CSIT_MS = "CSIT_MS"            ' ���� ����� CSIT �� Microsoft
Public EOL_CSIT_MS                          ' ��������� ������ ����� CSIT_MS

Public Const CSIT_MS_IDSF_COL = 1           ' ������� "IdSF" - Id ��������������������
                                            '                    ����������� � SF
Public Const CSIT_MS_NAME_COL = 2           ' ������� "��������" - ��� �����������
Public Const CSIT_MS_ADDR_COL = 5           ' ������� "�����"
Public Const CSIT_MS_STAMP = "������ ����������� ����������� ���������"

'!!!!!!!!!!!!!!!!!!!!!!!!!!!! ���� ������ !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Type PostAddr               '�������� ��������� ������
    PostIndex As String         ' ������
    City As String              ' �����
    State As String             ' �������/�����/����
    Street As String            ' �����, ���, ������
    Country As String           ' ������
    ErrFlag As Boolean          '= TRUE - ������ ������ ��������������
End Type

Type SNatr                  '�������� SN Autodesk
    sN      As String           ' �������� �����
    Contr   As String           ' �������� Autodesk / Agreement
    AccN    As String           ' Customer Account Number, CSN �� PartnerCenter
    AccName As String           ' ��� ������� �� PartnerCenter
    C_Start As Date             ' ���� ������ ��������� ADSK
    C_End   As Date             ' ���� ��������� ��������� ADSK
    C_Status As String          ' ������ ���������: Active, Expired, Inactive
    CM_Name As String           ' ��� �� - �������� �� SN
    CM_Mail As String           ' email CM
    CM_Tel  As String           ' ������� ��
    Description As String       ' �������� �������� Autodesk
    Seats As Integer            ' ����� ��������/���������� ����
    Status As String            ' ������ SN: Registered, Upgrade...
    Deployment As String        ' ��� �������� ������� ��� ���������
    Sbs As Boolean              ' TRUE ���� �������� �������
    Release_Year As String      ' ��� ������ �������� ADSK
    Registered As Date          ' ���� ����������� SN
    ErrFlag As Boolean          '= TRUE - ������/�������� SN ��������������
End Type

'<><><><><><><> ���� ��������� �� ������� <><><><><><><><><><>
Public Const FATAL_ERR = 1  ' ������� ��������� � Log ����, �� ����� � ������������
Public Const TYPE_ERR = 2   ' ������� ��������� � Log ���� � �� �����
Public Const WARNING = 3    ' ������� �������������� ������ � Log ����


