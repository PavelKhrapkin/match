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
        /// F_MATCH = "match.xlsm" - ��� ����� ������ ���������� match
        /// </summary>
        public const string F_MATCH = "match.xlsm";
        /// <summary>
        /// F_1C = "1C.xlsx"    - ���� ������� 1C: ��������, ���������, ������ ��������
        /// </summary>
        public const string F_1C = "1C.xlsx";
        /// <summary>
        /// F_SFDC = "SFDC.xlsx"    - ���� ������� Salesforce.com
        /// </summary>
        public const string F_SFDC = "SFDC.xlsx";
        /// <summary>
        /// F_ADSK = "ADSK.xlsm"    - ���� ������� Autodesk
        /// </summary>
        public const string F_ADSK = "ADSK.xlsm";
        /// <summary>
        /// F_STOCK = "Stock.xlsx"  - ���� ������� �� ������ � �������� �������
        /// </summary>
        public const string F_STOCK = "Stock.xlsx";
        /// <summary>
        /// ��������� ���� ��� ������������� �����������
        /// </summary>
        public const string F_TMP = "W_TMP.xlsm";

        public const string dirDBs = "C:\\Users\\Pavel_Khrapkin\\Documents\\Pavel\\match\\matchDBs\\";

        //-----------��������� ������� ��������� -----------------
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
