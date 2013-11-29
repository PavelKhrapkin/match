/*-----------------------------------------------------------------------
 * Document -- класс Документов проекта match 3.0
 * 
 *  27.11.2013  П.Храпкин, А.Пасс
 *  
 * - 27.11.13 переписано с VBA TOCmatch на С#
 * -------------------------------------------
 * Document(Name)          - КОНСТРУКТОР возвращает ОБЪЕКТ Документ с именем Name
 * 
 */
using System;
using Box = System.Windows.Forms.MessageBox;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Lib = match.MatchLib;

namespace ExcelAddIn2
{
    public class Document
    {
        private static bool initializedTOC = false;
        public static Dictionary<string, Document> Documents = new Dictionary<string, Document>();   //коллекция Документов

        //        public static List<Document> Documents = new List<Document>();   //коллекция Документов уже открытых в match
 
        private string name;
        private bool isOpen = false;
        private string FileName;
        private string SheetN;
        private string MadeStep;
        private DateTime MadeTime;
        private ulong chkSum;
        private int EOLinTOC;
        private List<Stamp> stampList;
        public Excel.Range Body;

        public const string F_MATCH = "match.xlsm";
        private const string TOC = "TOCmatch";
        private const int TOC_DIRDBS_COL = 10;  //в первой строке в колонке TOC_DIRDBS_COL записан путь к dirDBs
        private const int TOC_LINE = 4;         //строка номер TOL_LINE таблицы ТОС отностися к самому этому документу.
//        private static int EOL_toc;             //число строк в ТОС. Определяется при инициализации ТОС и хранится в TOC_LINE
//        private int iTOC;                       //номер строки в ТОС - атрибут Документа по имени name
        private const string dirDBs = "C:\\Users\\Pavel_Khrapkin\\Documents\\Pavel\\match\\matchDBs\\";    //временно!!!

        private const string F_1C = "1C.xlsx";
        private const string F_SFDC = "SFDC.xlsx";
        private const string F_ADSK = "ADSK.xlsm";
        private const string F_STOCK = "Stock.xlsx";
        private const string F_TMP = "W_TMP.xlsm";

/*
        protected struct tocRow {
            private  DateTime stepTime;
            string nameDoc;

            public tocRow(DateTime time) {
                stepTime = time;
                nameDoc = "";
            }
        };
*/
 //       private Dictionary<string, List<Stamp>> toc = 

//        private List<tocRow> tocBody = new List<tocRow>();

        /*
                Body body;
                Header hdr;
                Summary smr;
                LeftCols lft;
                RightCols rght;
         */

        public Document(string nameIn)
        {
            // конструктор Документа

            /* вначале, если нужно, инициируем структуры ТОС - Содержания или таблицы Документов */
            if (!initializedTOC)
            {
                Excel.Workbook db_match = fileOpen(F_MATCH);
                Excel.Worksheet wholeSheet = db_match.Worksheets[TOC];
                Excel.Range tocRng = wholeSheet.Range["5:" + Lib.EOL(wholeSheet)];
                bool insideTOCrw = true;
                foreach (Excel.ListRow rw in tocRng) {
                    if (insideTOCrw) {
                        MadeTime    = rw.Range["A1"];
                        name        = rw.Range["B1"];
                        EOLinTOC    = rw.Range["C1"];
    //                    MyCol        = rw.Range["D1"];
    //                    ResLines     = rw.Range["E1"];
                        MadeStep    = rw.Range["F1"];
    //                    Period    = rw.Range["G1"];
                        FileName    = rw.Range["H1"];
                        SheetN      = rw.Range["I1"];
                    }
                    //insideTOCrw = Stamp.initStamp(rw.Range["J1:M1,O1]");
                    stampList.Add(Stamp.initStamp(rw.Range["J1:M1"]));
                    insideTOCrw = String.IsNullOrEmpty(rw.Range["O1"]);
                }

                if (dirDBs != (string)db_match.Worksheets[TOC].cells[1, TOC_DIRDBS_COL].Value2)
                {
                    Box.Show("Файл '" + F_MATCH + "' загружен из необычного места!");
                    // переустановка match -- будем делать потом
                }
//                name = TOC;
//                FileName = F_MATCH;
//                SheetN = TOC;
//                Documents.Add(TOC, this);
                initializedTOC = true;
            }
            //                WrTOC(TOC);    /* WrTOC - метод, записывающий данные из приложения в лист TOCmatch - напишем позже */             }

            stampList = new List<Stamp>();  //каждый документ ссыылается на цепочку сигнатор-Штамп

            /* находим Документ name в ТОС проверяя его сигнатуры то есть Штамп */

<<<<<<< .mine
            Document doc = getDoc(nameIn);
            // не дописано
=======
            Document doc = null;
            foreach (KeyValuePair<string, Document> entry in OpenDocs) {
                if (entry.Value.Name == nameIn) {
                    doc = entry.Value;          // найден документ с именем nameIn
                    break;
                }
            }

>>>>>>> .r642
        }

        public static Document getDoc(string name)
        {
            return (Documents.ContainsKey(name)) ? Documents[name] : new Document(name);
        }

        public static Document loadDoc(string name, Excel.Workbook wb)
        {
            // загрузка в match нового документа
            // 27.11.13 -- еще не дописано

//            bool found = false;
            for (int i = 1; i <= Documents.Count; i++)
            {
                //                if (CheckStamp(Documents[i].Stamp)) {
            }
            return Documents[name];
        }

        public bool isDocOpen(string name) { return (Documents.ContainsKey(name)); }

        public static string recognizeDoc(Excel.Workbook wb) {
//            foreach (tocRow  rw in tocBody) {
//                if (tocRow.checkStamp(wb, tocBody.Stamp)) { return tocBody.name; }
//            }
            return null;
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }


        private static Excel.Workbook fileOpen(string name) {
            Excel.Application app = new Excel.Application();
            //            Microsoft.Office.Interop.Excel.Workbook wb;
            Excel.Workbook wb;
            try {
                wb = app.Workbooks.Open(dirDBs + name);
                return wb;
            } catch (Exception ex) {
                System.Windows.Forms.MessageBox.Show("Ошибка> " + ex.Message
                    + "\n не открыт файл '" + dirDBs + name + "'");
                return null;
            }
        }

        protected class Stamp {
            protected string signature; // проверяемый текст Штампа - сигнатура
            protected char typeStamp;   // '=' - точное соответствие сигнатуры; 'I' - "текст включает.."
            protected int rw, col;      // позиция сигнатуры в проверяемом Документе
            protected int or_rw, or_col;  // альтернативные позиции Штампа

            protected Stamp() {
            }

            public static Stamp initStamp(Excel.Range rng) {
                //
                return new Stamp(); 
            }

            public bool checkStamp(Excel.Workbook wb, Stamp stamp) {
                return false;
            }
        }    // конец класса Stamp      
    }    // конец класса Document
}