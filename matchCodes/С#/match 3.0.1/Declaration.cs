using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace match.Declaration
{
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

        public const string dirDBs = "C:\\Users\\Pavel_Khrapkin\\Documents\\Pavel\\match\\matchDBs\\";

        //-----------константы таблицы Процессов -----------------
        public const string STEP_TIME =     "A1";
        public const string STEP_NAME =     "B1";
        public const string STEP_PREV =     "C1";
        public const string STEP_COMMENT =  "D1";
        public const string STEP_DONE =     "E1";
        public const string STEP_PARAM =    "F1";
        public const string STEP_PARAMS =   "F1:J1";
        public const string STEP_DOCS =     "K1:O1";

        public const string PROC_NAME = STEP_PARAM;
        public const string PROC_DOCUMENTS = "K1:O1";
    }
}
