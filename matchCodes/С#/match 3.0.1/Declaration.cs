using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace match.Declaration
{
    /// <summary>
    /// Declaration -- определения глобальных констант
    /// </summary>
    /// <journal> 20.12.2013
    /// 7.1.2014    - добавлена секция констант Шаблонов
    /// </journal>
    class Declaration
    {
        /// <summary>
        /// F_MATCH = "match.xlsm" - имя файла таблиц приложения match
        /// </summary>
        public const string F_MATCH = "match.xlsm";
        /// <summary>
        /// F_1C = "1C.xlsx"    - файл отчетов 1C: Платежей, Договоров, Списка клиентов
        /// </summary>
        public const string F_1C = "1C.xlsx";
        /// <summary>
        /// F_SFDC = "SFDC.xlsx"    - файл отчетов Salesforce.com
        /// </summary>
        public const string F_SFDC = "SFDC.xlsx";
        /// <summary>
        /// F_ADSK = "ADSK.xlsm"    - файл отчетов Autodesk
        /// </summary>
        public const string F_ADSK = "ADSK.xlsm";
        /// <summary>
        /// F_STOCK = "Stock.xlsx"  - файл отчетов по Складу и Провозка Заказов
        /// </summary>
        public const string F_STOCK = "Stock.xlsx";
        /// <summary>
        /// временный файл для промежуточных результатов
        /// </summary>
        public const string F_TMP = "W_TMP.xlsm";

        public const string DIR_DBS = "ENV_DIR_DBS";

        //--------- листы match.xlsm --------------
        public const string WE = "We";
        public const string HEADER = "Header";
        public const string WP_PROTOTYPE = "WP_Prototype";

        //-----------константы таблицы Документов -----------------
        public const int DOC_TIME        = 1; // дата и время последнего изменения Документа
        public const int DOC_NAME        = 2; // имя Документа
        public const int DOC_EOL         = 3; // EOL Документа
        public const int DOC_MYCOL       = 4; // количество колонок, дабавляемых слева в Документ в loadDoc
        public const int DOC_RESLINES    = 5; // число строк пятки
        public const int DOC_MADESTEP    = 6; // последний выполненный Шаг
        public const int DOC_PERIOD      = 7; // периодичность работы с Документом в днях
        public const int DOC_FILE        = 8; // файл match, содержащий Документ
        public const int DOC_SHEET       = 9; // лист Документа
        public const int DOC_STMPTXT     = 10; // текст Штампа
        public const int DOC_STMPTYPE    = 11; // тип Штампа
        public const int DOC_STMPROW     = 12; // строка Штампа
        public const int DOC_STMPCOL     = 13; // колонка Штампа
        public const int DOC_CREATED     = 14; // дата создания Документа
        public const int DOC_PATTERN     = 16; // основной шаблон Документа
        public const int DOC_SUMMARY_PATTERN = 17; // шаблон пятки
        public const int DOC_LOADER      = 20; // Loader Документа

        //-----------константы таблицы Процессов -----------------
        public const string PROCESS = "Process";

        public const string STEP_TIME   = "A1";
        public const string STEP_NAME   = "B1";
        public const string STEP_PREV   = "C1";
        public const string STEP_COMMENT = "D1";
        public const string STEP_DONE   = "E1";
        public const string STEP_PARAM  = "F1";
        public const string STEP_PARAMS = "F1:J1";
        public const string STEP_DOCS   = "K1:O1";

        public const string PROC_NAME = STEP_PARAM;
        public const string PROC_DOCUMENTS = "K1:O1";

        //-----------константы Шаблонов -----------------
        public const string PTRN_HDR    = "A1";     // заголовки колонок   
        public const string PTRN_WIDTH  = "A3";     // ширина колонок
        public const int PTRN_FETCH = 6;    // Fetch запрос

        public const string PTRN_COPYHDR = "CopyHdr"; // указание копировать заголовок из Шаблона
    }
}