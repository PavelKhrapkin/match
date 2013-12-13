/*-----------------------------------------------------------------------
 * Document -- класс Документов проекта match 3.0
 * 
 *  14.12.2013  П.Храпкин, А.Пасс
 *  
 * - 10.12.13 переписано с VBA TOCmatch на С#
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
        private Stamp stamp;        //каждый документ ссылается на цепочку сигнатур или Штамп
        private DateTime creationDate;  // дата создания Документа
        private string Loader;
        private string BodyPtrn;
        private string SummPtrn;
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
                    Documents.Add(docName, doc);

                    // построить Range, включающий все штампы документа
                    int j;
                    for (j = i + 1; j <= tocRng.Rows.Count
                            && (String.IsNullOrEmpty(tocRng.Range["B" + j].Value2)); j++) ;
                    doc.stamp = new Stamp(tocRng.Range["J" + i + ":M" + --j]);

                    doc.creationDate = DateTime.FromOADate(rw.Range["N1"].Value2);

                    doc.BodyPtrn = rw.Range["P1"].Value2;
                    doc.SummPtrn = rw.Range["Q1"].Value2;
                    doc.Loader = rw.Range["T1"].Value2;
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

            Document doc = Documents[name];
            if (!doc.isOpen)
            {
                Excel.Worksheet Sh = fileOpen(doc.FileName).Worksheets[doc.SheetN];
// дальше распознавание частичное или полное одновление старого файла
// потом из wb переносим данные в старый файл
// а в конце запускаем Loader 
            }
            return doc;
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
    // проверки Штампов для Документа надо будет перенести в метод checkStampdocFailed
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
            int ee = Lib.EOL(wholeSheet);
            Excel.Range rng = wholeSheet.Range["1:" + Lib.EOL(wholeSheet).ToString()];

            // ищем подходящий документ в TOCmatch
            foreach (var doc in Documents)
            {
                // у F_SFDC штамп находится в конце документа
                int shiftToEol = (doc.Value.FileName == F_SFDC) ? rng.Rows.Count - 4 : 0;

                // цикл по штампам документа (сигнатурам) - все должны удовлетвориться
                bool allStampsOK = true;
                foreach (var stmp in doc.Value.stamp.stamps) {
                    // цикл по позициям для сигнатуры - хотя бы одна должна удовлетвориться
                    bool signOK = false;
                    foreach (var pos in stmp.stampPosition) {
                        // В ЭТОМ МЕСТЕ НЕ РАЗОБРАЛСЯ С ДОСТУПОМ К Cells
                        if (rng.Cells[pos[0] + shiftToEol, pos[1]].Value2 == stmp.signature) {
                            signOK = true;
                            break;
                        }
                    }   // конец цикла по позициям
                    if (!signOK) {
                        allStampsOK = false;
                        break;      // прерываем цикл по сигнатурам; к следующему документу
                    }
                }   // конец цикла по сигнатурам
                if (allStampsOK) return doc.Value.name;
            }       // конец цикла по документам
            return null;        // ничего не нашли
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
/*
 * Класс, описывающий все штампы документа 
 */
        private class Stamp
        {
            public List<OneStamp> stamps = new List<OneStamp>();
            /*
             * Конструктор. 
             *  rng - range, включающий колонки с J по М для всех строк, описывающих документ.
             */
            public Stamp(Excel.Range rng)
            {       // цикл
                if ((char)rng.Range["B1"].Value2[0] != 'N')
                {
                    for (int i = 1; i <= rng.Rows.Count; i++) stamps.Add(new OneStamp(rng.Rows[i]));
                }
            }
        }

    //           struct oneStamp {
    //               string signature
/*
* Класс, описывающий штамп документа (с вариантами позиций, заданными в одной стрке TOCmatch)
*/
        public class OneStamp
        {
            public string signature;  // проверяемый текст Штампа - сигнатура
            private char typeStamp;   // '=' - точное соответствие сигнатуры; 'I' - "текст включает.."
            public List<int[]> stampPosition = new List<int[]>();   // альтернативные позиции сигнатур Штампов
            /*
            * Конструктор
            *  rng - range, включающий одну строку штампа (т.е. сигнатуру)
            *  
            * примеры: {[1, "1, 6"]} --> [1,1] или [1,6]
            *  .. {["4,1", "2,3"]} --> [4,2]/[4,3]/[1,2]/[1,3]
            */
            public OneStamp(Excel.Range rng)
            {
                signature = rng.Range["A1"].Value2; 
                typeStamp = rng.Range["B1"].Value2[0];

                List<int> rw = intListFrCell("C1", rng);
                List<int> col = intListFrCell("D1", rng);
                // декартово произведение множеств rw и col
                rw.ForEach(r => col.ForEach(c => stampPosition.Add(new int[] { r, c })));
            }

            private List<int> intListFrCell(string coord, Excel.Range rng)
            {
                return Lib.ToIntList(rng.Range[coord].Value2.ToString(), ',');
            }

        }   // конец класса OneStamp
    }    // конец класса Document
}