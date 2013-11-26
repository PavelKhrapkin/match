/*-----------------------------------------------------------------------
 * Document -- класс Документов проекта match 3.0
 * 
 *  26.11.2013  П.Храпкин, А.Пасс
 *  
 * - 26.11.13 переписано с VBA TOCmatch на С#
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

        private string name;
        private string FileName;
        private string SheetN;
        private string MadeStep;
        private DateTime MadeTime;
        private ulong chkSum;
        private int EOLinTOC;
        //             Stamp;
        public Excel.Range Body;

        public const string F_MATCH = "match.xlsm";
        private const string TOC = "TOCmatch";
        private const int TOC_DIRDBS_COL = 10;  //в первой строке в колонке TOC_DIRDBS_COL записан путь к dirDBs
        private const int TOC_LINE = 4;         //строка номер TOL_LINE таблицы ТОС отностися к самому этому документу.
        private static int EOL_toc;             //число строк в ТОС. Определяется при инициализации ТОС и хранится в TOC_LINE
        private int iTOC;                       //номер строки в ТОС - атрибут Документа по имени name
        private const string dirDBs = "C:\\Users\\Pavel_Khrapkin\\Documents\\Pavel\\match\\matchDBs\\";    //временно!!!

        private const string F_1C = "1C.xlsx";
        private const string F_SFDC = "SFDC.xlsx";
        private const string F_ADSK = "ADSK.xlsm";
        private const string F_STOCK = "Stock.xlsx";
        private const string F_TMP = "W_TMP.xlsm";

//        public static List<Document> OpenDocs = new List<Document>();   //коллекция Документов уже открытых в match
        public Dictionary<string, Document> OpenDocs = new Dictionary<string, Document>();   //коллекция Документов уже открытых в match
        /*
                Stamp stamp;
 
                Body body;
                Header hdr;
                Summary smr;
                LeftCols lft;
                RightCols rght;
         */

        /*
         * Конструктор Документа
         */
        public Document(string nameIn)
        {
            /* вначале, если нужно, инициируем структуры ТОС - Содержания или таблицы Документов */
            if (!isDocOpen(TOC))
            {
                Excel.Workbook db_match = fileOpen(F_MATCH);
                Excel.Worksheet wholeSheet = db_match.Worksheets[TOC];
                EOLinTOC = Lib.EOL(wholeSheet);
                Body = db_match.Worksheets[TOC].get_Range("A4:ZZ" + EOLinTOC);
                if (dirDBs != (string)db_match.Worksheets[TOC].cells[1, TOC_DIRDBS_COL].Value2)
                {
                    Box.Show("Файл '" + F_MATCH + "' загружен из необычного места!");
                    // переустановка match -- будем делать потом
                }
                name = TOC;
                FileName = F_MATCH;
                SheetN = TOC;

//              OpenDocs.Add(this);
                OpenDocs.Add(TOC, this);            }
            //                WrTOC(TOC);    /* WrTOC - метод, записывающий данные из приложения в лист TOCmatch - напишем позже */             }

            /* находим Документ name в ТОС проверяя его сигнатуры то есть Штамп */
            for (iTOC = 1; iTOC <= OpenDocs[1].Body.Rows.Count; iTOC++)
            {
                //                   if (nameIn == Body
                //               if TOC.
            }

        }

        public static Document getDoc(string name)
        {
/*
            foreach (Document Doc in OpenDocs)
            {
                if (Doc.Name == name) return Doc;
            }
            return new Document(name);
 */
            return (OpenDocs.ContainsKey(name)) ? OpenDocs[name] : new Document(name);
        }

        public static bool isDocOpen(string name)
        {
/*
            foreach (Document Doc in OpenDocs)
            {
                if (Doc.Name == name) return true;
            }
            return false;
 */
            return (OpenDocs.ContainsKey(name));
        }
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        public bool CheckStamp()
        {
            return false;
        }

        private static Excel.Workbook fileOpen(string name)
        {

            Excel.Application app = new Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb;
            try
            {
                wb = app.Workbooks.Open(dirDBs + name);
                return wb;
            }
            catch
            {
                return null;
            }
        }    // конец класса Document
    }
}
