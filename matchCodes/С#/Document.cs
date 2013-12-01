//*-----------------------------------------------------------------------
 * Document -- класс Документов проекта match 3.0
 * 
 *  1.12.2013  П.Храпкин, А.Пасс
 *  
 * - 1.12.13 переписано с VBA TOCmatch на С#
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
        private static Dictionary<string, Document> Documents = new Dictionary<string, Document>();   //коллекция Документов
 
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
        private const string dirDBs = "C:\\Users\\Pavel_Khrapkin\\Documents\\Pavel\\match\\matchDBs\\";    //временно!!!

        private const string F_1C = "1C.xlsx";
        private const string F_SFDC = "SFDC.xlsx";
        private const string F_ADSK = "ADSK.xlsm";
        private const string F_STOCK = "Stock.xlsx";
        private const string F_TMP = "W_TMP.xlsm";

        public Document(string nameIn)
        {
            // конструктор Документа

            /* вначале, если нужно, инициируем структуры ТОС - Содержания или таблицы Документов */
            if (!initializedTOC)
            {
                Excel.Workbook db_match = fileOpen(F_MATCH);
                Excel.Worksheet wholeSheet = db_match.Worksheets[TOC];
                Excel.Range tocRng = wholeSheet.Range["5:" + Lib.EOL(wholeSheet)];
                foreach (Excel.ListRow rw in tocRng) {
                    string docName = rw.Range["B1"];
                    if (!String.IsNullOrEmpty(docName)) {
                        MadeTime = rw.Range["A1"];
                        name = docName;
                        EOLinTOC    = rw.Range["C1"];
    //                    MyCol        = rw.Range["D1"];
    //                    ResLines     = rw.Range["E1"];
                        MadeStep    = rw.Range["F1"];
    //                    Period    = rw.Range["G1"];
                        FileName    = rw.Range["H1"];
                        SheetN      = rw.Range["I1"];
                    }
                    stampList.Add(new Stamp(rw.Range["J1:M1"]));
                }

                if (dirDBs != (string)db_match.Worksheets[TOC].cells[1, TOC_DIRDBS_COL].Value2)
                {
                    Box.Show("Файл '" + F_MATCH + "' загружен из необычного места!");
                    // переустановка match -- будем делать потом
                }
                initializedTOC = true;
            }
            //                WrTOC(TOC);    /* WrTOC - метод, записывающий данные из приложения в лист TOCmatch - напишем позже */             }

            stampList = new List<Stamp>();  //каждый документ ссыылается на цепочку сигнатор-Штамп

            /* находим Документ name в ТОС проверяя его сигнатуры то есть Штамп */

 //         Document doc = getDoc(nameIn);
            // не дописано
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
            Document emp = new Document("");
            Excel.Worksheet wholeSheet = wb.Worksheets[1];
            Excel.Range rng = wholeSheet.Range["1:" + Lib.EOL(wholeSheet)];

            foreach (var doc in Documents)
            {
                foreach (Stamp stmp in doc.Value.stampList)
                {
                    foreach (int[] pos in stmp.stampPosition)
                    {
                        if (rng.Cells[pos[0], pos[1]].Value2 != stmp.signature) break;
                    }
                    return doc.Value.name;
                }
            }
            return null;
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
            public string signature; // проверяемый текст Штампа - сигнатура
            protected char typeStamp;   // '=' - точное соответствие сигнатуры; 'I' - "текст включает.."

            // альтернативная позиция Штампа, если есть, сохраняется во второй компоненте массива
            public List<int[]> stampPosition;      // позиция сигнатуры в проверяемом Документе

            public Stamp(Excel.Range rng)
            {
                signature = rng.Cells[1, 1].value;                          // {[1, "1, 6"]} --> [1,1] или [1,6]
                typeStamp = rng.Cells[1, 2].value;
                List<int> rw = Lib.ToIntList(rng.Cells[1, 3].value, ',');
                List<int> col = Lib.ToIntList(rng.Cells[1, 4].value, ',');

                for (int j = col.Count + 1; j <= rw.Count; j++) col.Add(rw[j]);
                for (int j = rw.Count + 1; j <= col.Count; j++) rw.Add(col[j]);
                for (int j = rw.Count + 1; j <= col.Count; j++) {
                    int[] x = {rw[j],col[j]};
                    stampPosition.Add(x);        
                }
           }
            public bool checkStamp(Excel.Workbook wb, Stamp stamp) {
                return false;
            }
        }    // конец класса Stamp      
    }    // конец класса Document
}