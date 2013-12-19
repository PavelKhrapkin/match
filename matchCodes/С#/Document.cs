/*-----------------------------------------------------------------------
 * Document -- класс Документов проекта match 3.0
 * 
 *  19.12.2013  П.Храпкин, А.Пасс
 *  
 * - 19.12.13 переписано с VBA TOCmatch на С#
 * -------------------------------------------
 * Document(name)       - КОНСТРУКТОР возвращает ОБЪЕКТ Документ с именем name
 * loadDoc(name, wb)    - загружает Документ name или его обновления из файла wb
 * getDoc(name)         - возвращает Документ с именем name; при необходимости - открывает его
 * isDocOpen(name)      - проверяет, что Документ name открыт
 * recognizeDoc(wb)     - распознает первый лист файла wb по таблице Штампов
 * Check(rng,stampList)       - проверка Штампов stampList в Range rng
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
    /// <summary>
    /// класс Document содержит таблицу параметров всех Документов, известных приложению match
    /// </summary>
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
        private bool isPartialLoadAllowed;
        private string BodyPtrn;
        private string SummPtrn;
        public Excel.Range Body;
        public Excel.Range Summary;

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

        private const string TOC = "TOCmatch";
        private const int TOC_DIRDBS_COL = 10;  //в первой строке в колонке TOC_DIRDBS_COL записан путь к dirDBs
        private const int TOC_LINE = 4;         //строка номер TOL_LINE таблицы ТОС отностися к самому этому документу.
//        private const string dirDBs = "C:\\Users\\Pavel_Khrapkin\\Documents\\Pavel\\match\\matchDBs\\";    //временно!!!

        static Document()
        {
            Document doc = null;
            Excel.Workbook db_match = FileOpenEvent.fileOpen(F_MATCH);
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
                    bool isSF = doc.FileName == F_SFDC;
                    doc.stamp = new Stamp(tocRng.Range["J" + i + ":M" + --j], isSF);

                    doc.creationDate = DateTime.FromOADate(rw.Range["N1"].Value2);

                    doc.BodyPtrn = rw.Range["P1"].Value2;
                    doc.SummPtrn = rw.Range["Q1"].Value2;
                    doc.Loader = rw.Range["T1"].Value2;

                    // флаг, разрешающий частичное обновление Документа пока прописан хардкодом
                    switch (docName)
                    {
                        case "Платежи":
                        case "Договоры": doc.isPartialLoadAllowed = true;
                            break;
                        default: doc.isPartialLoadAllowed = false;
                            break;
                    }
                }
            }
            //if (dirDBs != (string)db_match.Worksheets[TOC].cells[1, TOC_DIRDBS_COL].Value2)
            //{
            //    Box.Show("Файл '" + F_MATCH + "' загружен из необычного места!");
            //    // переустановка match -- будем делать потом
            //}
        }
        /// <summary>
        /// loadDoc(name, wb)   - загрузка содержимого Документа name из файла wb
        /// </summary>
        /// <param name="name"></param>
        /// <param name="wb"></param>
        /// <returns>Document   - при необходимости читает name из файла в match и сливает его с данными в wb</returns>
        /// <journal> Не дописано
        /// 15.12.2013 - взаимодействие с getDoc(name)
        /// </journal>
        public static Document loadDoc(string name, Excel.Workbook wb)
        {
            Document doc = getDoc(name);
            Excel.Workbook wb_sf = FileOpenEvent.fileOpen(doc.FileName);
            Excel.Worksheet Sh = wb_sf.Worksheets[doc.SheetN];
//            Excel.Worksheet Sh = fileOpen(doc.FileName).Worksheets[doc.SheetN];
            if (doc.isPartialLoadAllowed)
            {
// дальше распознавание частичное или полное обновление прежнего Документа
// здесь только если частичное, то есть потом будет выполняться Merge
            }
            wb.Worksheets[1].Name = "TMP";
            wb.Worksheets[1].Move(Sh);
// потом из wb переносим данные в старый файл
// а в конце запускаем Loader
            return doc;
        }
        /// <summary>
        /// getDoc(name)            - извлечение Документа name. Если еще не прочтен - из файла
        /// </summary>
        /// <param name="name">имя извлекаемого документа</param>
        /// <returns>Document</returns>
        /// <journal> Не дописано
        /// 15.12.2013 - чтение из файла, формирование Range Body и Summary
        /// </journal>
        public static Document getDoc(string name)
        {
            try
            {
                Document doc = Documents[name];
                if (!doc.isOpen)
                {
                    // загрузка Документа из файла
                    // надо проверить Штампы открытого документа в Sh, его EOL,
                    // определить его Range Body и Summary                }
                    //!!!!!                if (!Check(rng, doc.stampList))
// это неправильно!!                    Excel.Range rng = (doc.FileName == F_SFDC) ? doc.Summary : doc.Body;
                }
                return doc;
            }
            catch
            {
                // надо проверить, что Document name не существует
                // в случае, если существует, но не удалось прочитать - создать событие FATAL_ERR
                return null;
            }
        }
        /// <summary>
        /// isDocOpen(name)     - проверяет, что Документ name открыт и доступен
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <juornal> 10.12.2013
        /// </juornal> 
        public bool isDocOpen(string name) { return (Documents.ContainsKey(name)); }
        /// <summary>
        /// recognizeDoc(wb)        - распознавание Документа в Листе[1] wb
        /// </summary>
        /// <param name="wb"></param>
        /// <returns>имя распознанного документа или null, если Документ не распознан</returns>
        /// <journal> 14.12.2013
        /// 16.12.13 (ПХ) переписано распознавание с учетом if( is_wbSF(wb) )
        /// </journal>
        public static string recognizeDoc(Excel.Workbook wb) {
            Excel.Worksheet wholeSheet = wb.Worksheets[1];
            Excel.Range rng = wholeSheet.Range["1:" + Lib.EOL(wholeSheet).ToString()];

            Stamp stmpSF = Documents["SFDC"].stamp;
            bool is_wbSF = Stamp.Check(rng, stmpSF);
            // ищем подходящий документ в TOCmatch
            foreach (var doc in Documents)
            {
                if (is_wbSF && (doc.Value.FileName != F_SFDC)) continue;
                if (doc.Value.name == "SFDC" || doc.Value.name == "Process") continue;
                if (Stamp.Check(rng, doc.Value.stamp)) return doc.Value.name;
            }       // конец цикла по документам
            return null;        // ничего не нашли
        }

        /// <summary>
        /// Класс Stamp, описывающий все штампы документа
        /// </summary>    
        private class Stamp
        {
            public List<OneStamp> stamps = new List<OneStamp>();
            /*
             * Конструктор. 
             *  rng - range, включающий колонки с J по М для всех строк, описывающих документ.
             */
            public Stamp(Excel.Range rng, bool isSF)
            {       // цикл
                if ((char)rng.Range["B1"].Value2[0] != 'N')
                {
                    for (int i = 1; i <= rng.Rows.Count; i++) stamps.Add(new OneStamp(rng.Rows[i], isSF));
                }
            }
            /// <summary>
            /// Check(rng, stmp)        - проверка, что Range rng соответствует цепочке Штампов в stmp
            /// </summary>
            /// <param name="rng">Range rng - проверяемый Документ</param>
            /// <param name="stmp">Stamp stmp   - цепочка Штампов, соответствуюших данному Документу</param>
            /// <returns>true, если результат проверки положительный, иначе false</returns>
            /// <journal> 12.12.13
            /// 16.12.13 (ПХ) перенес в класс Stamp и переписал
            /// </journal>
            public static bool Check(Excel.Range rng, Stamp stmp)
            {
                foreach (OneStamp st in stmp.stamps)
                    if (!OneStamp.Check(rng, st)) return false;
                return true;
            }
        }

        /// <summary>
        /// Класс, описывающий штамп документа (с вариантами позиций, заданными в одной стрке TOCmatch)
        /// </summary>
        public class OneStamp
        {
            private string signature;  // проверяемый текст Штампа - сигнатура
            private char typeStamp;   // '=' - точное соответствие сигнатуры; 'I' - "текст включает.."
            private List<int[]> stampPosition = new List<int[]>();   // альтернативные позиции сигнатур Штампов
            private bool _isSF;
 
            /// <summary>
            /// Конструктор OneStanp(rng, isSF)
            /// </summary>
            /// <param name="rng">rng - range, включающий одну строку штампа (т.е. сигнатуру)</param>
            /// <param name="isSF">isSF</param>
            /// <example>
            /// примеры: {[1, "1, 6"]} --> [1,1] или [1,6]
            ///  .. {["4,1", "2,3"]} --> [4,2]/[4,3]/[1,2]/[1,3]
            /// </example>
            /// <journal> 12.12.2013 (AP)
            /// 16.12.13 (ПХ) добавлен параметр isSF - добавляется в структуру Штампа
            /// </journal>
            public OneStamp(Excel.Range rng, bool isSF)
            {
                signature = rng.Range["A1"].Value2; 
                typeStamp = rng.Range["B1"].Value2[0];
                _isSF = isSF;

                List<int> rw = intListFrCell("C1", rng);
                List<int> col = intListFrCell("D1", rng);
                // декартово произведение множеств rw и col
                rw.ForEach(r => col.ForEach(c => stampPosition.Add(new int[] { r, c })));
            }
            /// <summary>
            /// Check(rng, stmp)        - проверка сигнатуры Штампа stmp в rng для его всех допустимых позиций
            /// </summary>
            /// <param name="rng"></param>
            /// <param name="stmp"></param>
            /// <returns></returns>
            public static bool Check(Excel.Range rng, OneStamp stmp)
            {
                int shiftToEol = (stmp._isSF) ? rng.Rows.Count - 6 : 0;
                string sig = stmp.signature.ToLower();
                foreach (var pos in stmp.stampPosition)
                {
                    var x = rng.Cells[pos[0] + shiftToEol, pos[1]].Value2;
                    if (x == null) continue;
                    string strToCheck = x.ToLower();

                    if (stmp.typeStamp == '=')
                    {
                        if (strToCheck == sig) return true;
                    }
                    else
                    {
                        if (strToCheck.Contains(sig)) return true;
                    }
                }
                return false;
            }

            private List<int> intListFrCell(string coord, Excel.Range rng)
            {
                return Lib.ToIntList(rng.Range[coord].Value2.ToString(), ',');
            }

        }   // конец класса OneStamp
    }    // конец класса Document
}