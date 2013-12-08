/*-----------------------------------------------------------------------
 * Document -- класс Документов проекта match 3.0
 * 
 *  9.12.2013  П.Храпкин, А.Пасс
 *  
 * - 9.12.13 переписано с VBA TOCmatch на С#
 * -------------------------------------------
 * Document(name)       - КОНСТРУКТОР возвращает ОБЪЕКТ Документ с именем name
 * loadDoc(name, wb)    - загружает Документ name или его обновления из файла wb
 * getDoc(name)         - возвращает Документ с именем name; при необходимости - открывает его
 * isDocOpen(name)      - проверяет, что Документ name открыт
 * recognizeDoc(wb)     - распознает первый лист файла wb по таблице Штампов
 * checkStamp(rng,stampList)       - проверка Штампов stampList в Range rng
 * 
 * внутренний класс Stamp предназначен для заполнения списков Штампов
 * каждый Штамп содержит сигнатуру, то есть проверяемый текст, и пар координат - его положений
 * Stamp(Range rng)     - разбирает rng, помещая из таблицы TOCmatch Штамп в List Штампов в Документе
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
        public Excel.Range Summary;

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

        static Document()
        {
            Document doc = null;
            Excel.Workbook db_match = fileOpen(F_MATCH);
            Excel.Worksheet wholeSheet = db_match.Worksheets[TOC];
            Excel.Range tocRng = wholeSheet.Range["5:" + Lib.EOL(wholeSheet)];

            for (int i = 1; i <= tocRng.Rows.Count; i++)
            {
                Excel.Range rw = tocRng.Rows[i];
 
                string docName = rw.Range["B1"].Value2;
                if (!String.IsNullOrEmpty(docName))
                {
                    doc = new Document();
                    doc.MadeTime = DateTime.FromOADate(rw.Range["A1"].Value2);
                    doc.name = docName;
                    string tx = rw.Range["C1"].Value2.ToString();
                    doc.EOLinTOC = String.IsNullOrEmpty(tx) ? 0 : Convert.ToInt32(tx);
                    //                    MyCol        = rw.Range["D1"].Value2;
                    //                    ResLines     = rw.Range["E1"].Value2;
                    doc.MadeStep = rw.Range["F1"].Value2;
                    //                    Period    = rw.Range["G1"].Value2;
                    doc.FileName = rw.Range["H1"].Value2;
                    doc.SheetN = rw.Range["I1"].Value2;
                    Documents.Add(docName,doc);
                    doc.stampList = new List<Stamp>();  //каждый документ ссыылается на цепочку сигнатор-Штамп
                }
                Excel.Range rwStamp = rw.Range["J1:M1"];
                if (rwStamp.Cells[1, 2].value.ToString()[0] != 'N') // тип Штампа 'N' - проверка Штампа не нужна
                {
                    doc.stampList.Add(new Stamp(rwStamp));
                }

            }
            if (dirDBs != (string)db_match.Worksheets[TOC].cells[1, TOC_DIRDBS_COL].Value2)
            {
                Box.Show("Файл '" + F_MATCH + "' загружен из необычного места!");
                // переустановка match -- будем делать потом
            }
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

        public Document getDoc(string name)
        {
            try
            {
                Document doc = Documents[name];
                if (!doc.isOpen)
                {
                    // загрузка Документа из файла
                }
                Excel.Range rng = (doc.FileName == F_SFDC) ? doc.Summary : doc.Body;
//!!!!!                if (!checkStamp(rng, doc.stampList))
                {
                    // фатальная ошибка
                }
                return doc;
            }
            catch
            {
                return null;
            }
    // ДОПИСАТЬ!! Но вначале надо отладить recognizeDoc. При этом скорее всего внутреннюю часть цикла
    // проверки Штампов для Документа надо будет перенести в метод checkStamp
        }

        private bool checkStamp(Excel.Range rng, List<Stamp> list)
        {
            throw new NotImplementedException();
        }

        private bool checkStamp(Excel.Range rng, Document stampList)
        {
            return false;
        }
 
        public bool isDocOpen(string name) { return (Documents.ContainsKey(name)); }

        public static string recognizeDoc(Excel.Workbook wb) {
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
            app.Visible = true;
            foreach (Excel.Workbook W in app.Workbooks)
            {
                if (W.Name == name)
                {
                    if (W.ActiveSheet.IsNullOrEmpty()) continue;
                    return W;
                }
            }
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
            public string signature;  // проверяемый текст Штампа - сигнатура
            private char typeStamp;   // '=' - точное соответствие сигнатуры; 'I' - "текст включает.."
            public List<int[]> stampPosition = new List<int[]>();   // альтернативные позиции сигнатур Штампов

            public Stamp(Excel.Range rng)
            {
                signature = rng.Cells[1, 1].value;  // примеры: {[1, "1, 6"]} --> [1,1] или [1,6]
                string str = rng.Cells[1, 2].value; // .. {["4,1", "2,3"]} --> [4,2]/[4,3]/[1,2]/[1,3]
                typeStamp = str[0];

                List<int> rw = Lib.ToIntList(rng.Cells[1, 3].value.ToString(), ',');
                List<int> col = Lib.ToIntList(rng.Cells[1, 4].value.ToString(), ',');

                for (int i = 0; i < rw.Count; i++)
                {
                    for (int j = 0; j < col.Count; j++)
                    {
                        int[] x = { rw[i], col[j] };
                        stampPosition.Add(x);
                    }
                }
            }
        }    // конец класса Stamp      
    }    // конец класса Document
}