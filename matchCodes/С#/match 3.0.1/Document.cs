/*-----------------------------------------------------------------------
 * Document -- класс Документов проекта match 3.0
 * 
 *  14.1.2014  П.Храпкин, А.Пасс
 *  
 * -------------------------------------------
 * Document(name)       - КОНСТРУКТОР возвращает ОБЪЕКТ Документ с именем name
 * loadDoc(name, wb)    - загружает Документ name или его обновления из файла wb, запускает Handler Документа
 * getDoc(name)         - возвращает Документ с именем name; при необходимости - открывает его
 * isDocChanged(name)      - проверяет, что Документ name открыт
 * recognizeDoc(wb)     - распознает первый лист файла wb по таблице Штампов
 * 
 * внутренний класс Stamp предназначен для заполнения списков Штампов
 * каждый Штамп содержит сигнатуру, то есть проверяемый текст, и пар координат - его положений
 * Stamp(Range rng)     - разбирает rng, помещая из таблицы TOCmatch Штамп в List Штампов в Документе
 * Check(rng,stampList) - проверка Штампов stampList в Range rng 
 */
using System;
//using Box = System.Windows.Forms.MessageBox;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using match.MyFile;
using Decl = match.Declaration.Declaration;
//using Mtrx = match.Matrix.Matrix;
using Mtr = match.Matrix.Matr;  //--
using Lib = match.Lib;
using Log = match.Lib.Log;
using Proc = match.Process.Process;

namespace match.Document
{
    /// <summary>
    /// класс Document содержит таблицу параметров всех Документов, известных приложению match
    /// </summary>
    public class Document
    {
        private static Dictionary<string, Document> Documents = new Dictionary<string, Document>();   //коллекция Документов

        public string name;
        private bool isOpen = false;
        public bool isChanged = false;
        private string FileName;
#if WB_PRIVATE
        //private Excel.Workbook wb;
        //public Excel.Workbook Wb { get; private set; }
#else
        public Excel.Workbook Wb;
#endif
        private string SheetN;
        public Excel.Worksheet Sheet;
        private string MadeStep;
        private DateTime MadeTime;
        private int periodDays;     // периодичность работы с Документом в днях
        private ulong chkSum;
        private int EOLinTOC;
        private List<int> ResLines; //число строк в пятке -- возможны альтернативные значения
        private Stamp stamp;        //каждый документ ссылается на цепочку сигнатур или Штамп
        private DateTime creationDate;  // дата создания Документа
        private string Loader;
        private string LastUpdateFromFile;
        private bool isPartialLoadAllowed;
        public int MyCol;           // количесто колонок, добавляемых слева в Документ в loadDoc
        public int usedColumns;     // общее кол-во использованных колонок в Body Документа
        public Mtr ptrn;
        public Mtr Body;
        public Mtr Summary;
        public Dictionary<string, Dictionary<string, string>> docDic = new Dictionary<string, Dictionary<string, string>>();

        private const string TOC = "TOCmatch";
        private const int TOC_DIRDBS_COL = 10;  //в первой строке в колонке TOC_DIRDBS_COL записан путь к dirDBs
        private const int TOC_LINE = 4;         //строка номер TOL_LINE таблицы ТОС отностися к самому этому документу.

        static Document()
        {
            Log.set("статический конструктор Документов");
            Document doc = null;
            Excel.Workbook db_match = FileOpenEvent.fileOpen(Decl.F_MATCH);
            Excel.Worksheet tocSheet = db_match.Worksheets[TOC];

            Excel.Worksheet hdrSht = db_match.Worksheets[Decl.HEADER]; // в этом листе шапки в именованных Range по всем Документам
            Excel.Worksheet wpPrototype = db_match.Worksheets[Decl.WP_PROTOTYPE]; // в этом листе ?

            Mtr mtr = new Mtr(tocSheet.UsedRange.get_Value());  //перенос данных из Excel в память

            int iEOL = mtr.iEOL();

            for (int i = 4; i <= iEOL; i++)
            {
                string docName = mtr.String(i, Decl.DOC_NAME);
                if (docName != "")
                {
                    doc = new Document();
                    doc.name     = docName;
                    if (doc.name == TOC)    // mtr относится только к TOCmatch, а не ко всем Документам  
                    {                       
                        doc.Body = mtr;
                        doc.Wb = db_match;
                        doc.Sheet = tocSheet;
                        doc.EOLinTOC = iEOL;
                        doc.isOpen = true;
                    }
                    Documents.Add(docName, doc);

                    doc.MadeTime = Lib.MatchLib.getDateTime(mtr.get(i, Decl.DOC_TIME));
                    doc.EOLinTOC = mtr.Int(i, Decl.DOC_EOL, "не распознан EOL в строке " + i);
                    doc.ResLines = Lib.MatchLib.ToIntList(mtr.String(i, Decl.DOC_RESLINES), '/');
                    doc.MyCol    = mtr.Int(i, Decl.DOC_MYCOL, "не распознан MyCol в строке " + i);
                    doc.MadeStep = mtr.String(i, Decl.DOC_MADESTEP);
                    doc.periodDays = mtr.Int(i, Decl.DOC_PERIOD);
                    doc.FileName = mtr.String(i, Decl.DOC_FILE);
                    doc.SheetN   = mtr.String(i, Decl.DOC_SHEET);
                    doc.creationDate = Lib.MatchLib.getDateTime(mtr.get(i, Decl.DOC_CREATED));
                    string ptrnName = mtr.String(i, Decl.DOC_PATTERN);
                    try {
                        if (ptrnName != "") doc.ptrn = new Mtr(hdrSht.Range[ptrnName].get_Value());
                    } catch {
                        if (ptrnName == "WP_Prototype") continue;
                        //DefinedNames rngInMatch =  db_match.DefinedNames;
                        //if (ptrnName == "WP_Prototype") {
                        //doc.ptrn = new Mtr(wpPrototype.Range[ptrnName].get_Value());
                        //}
                    } // затычка
        //            if (ptrnName != "") doc.ptrn = new Mtr(db_match.
                    int j;
                    for (j = i + 1; j <= iEOL && mtr.String(j, Decl.DOC_NAME) == ""; j++) ;
                    doc.stamp = new Stamp(i, j - 1);
                } //if docName !=""
            } // for по строкам TOC

 //                   try { doc.creationDate = Lib.MatchLib.getDateTime(Double.Parse(rw.Range[Decl.DOC_CREATED].Text)); }
 //                   catch { doc.creationDate = new DateTime(0); }

 //                   try { doc.ptrn = hdrSht.get_Range((string)rw.Range[Decl.DOC_PATTERN].Text); } catch { doc.ptrn = null; }
 //                   try { doc.SummaryPtrn = hdrSht.get_Range((string)rw.Range[Decl.DOC_SUMMARY_PATTERN].Text); } catch { doc.SummaryPtrn = null; }
 //                   doc.Loader = rw.Range[Decl.DOC_LOADER].Text;
 //                   // флаг, разрешающий частичное обновление Документа пока прописан хардкодом
 //                   switch (docName)
 //                   {
 //                       case "Платежи":
 //                       case "Договоры": doc.isPartialLoadAllowed = true;
 //                           break;
 //                       default: doc.isPartialLoadAllowed = false;
 //                           break;
 //                   }
 //               }
 //           }

            //-----------------------------------------------------------------
            // из коллекции Documents переносим произошедшие изменения в файл
            //            if (doc.Body.Range["A" + TOC_DIRDBS_COL].Value2 != Decl.dirDBs)
            {
                //    Box.Show("Файл '" + F_MATCH + "' загружен из необычного места!");
                //    // переустановка match -- будем делать потом
                doc.isChanged = true;
            }
//            doc.EOLinTOC = iEOL;
//PK            doc.Body.Range["C4"].Value2 = iEOL.ToString();
            doc.isChanged = true;   // TOCmatch сохраняем всегда. Возможно, это времянка
//            doc.isOpen = true;
            doc.saveDoc();
            Log.exit();
        }
        /// <summary>
        /// loadDoc(name, wb)   - загрузка содержимого Документа name из файла wb
        /// </summary>
        /// <param name="name"></param>
        /// <param name="wb"></param>
        /// <returns>Document   - при необходимости читает name из файла в match и сливает его с данными в wb</returns>
        /// <journal> Не дописано
        /// 15.12.2013 - взаимодействие с getDoc(name)
        /// 7.1.13 - выделяем в Документе Body и пятку посредством splitBodySummary
        /// </journal>
        public static Document loadDoc(string name, Excel.Workbook wb)
        {
            Log.set("loadDoc(" + name + ", " + wb.Name + ")");
            Document doc = getDoc(name);
            if (doc.isPartialLoadAllowed)
            {
                // Дальше распознавание частичное или полное обновление прежнего Документа.
                // Здесь только если частичное, то есть потом будет выполняться Merge
            }
            doc.LastUpdateFromFile = wb.Name;
            string oldRepName = "Old_" + doc.SheetN;
            try
            {
                wb.Worksheets[1].Name = "TMP";
                wb.Worksheets[1].Move(doc.Sheet);
                 // если Old_ уже есть, но еще не обработан - уничтожаем прежний "частичный" отчет
                if (match.MyFile.FileOpenEvent.sheetExists(doc.Wb, oldRepName))
                {
                    match.MyFile.FileOpenEvent.DisplayAlert(false);
                    doc.Wb.Worksheets[doc.name].Delete();
                    match.MyFile.FileOpenEvent.DisplayAlert(true);
                } else doc.Sheet.Name = "Old_" + doc.SheetN;
                doc.Wb.Worksheets["TMP"].Name = doc.SheetN;
            }
            catch
            {
                Log.FATAL("Не удалось перенести лист [1] из входного файла "
                    + doc.LastUpdateFromFile + " в Документ " + name);
            }
            doc.Sheet = doc.Wb.Worksheets[name];
            doc.splitBodySummary();
//PK            new Log("Во входном файле типа \"" + doc.name + "\" " + doc.Body.Rows.Count + " строк");

            // если есть --> запускаем Handler
            if (doc.Loader != null) Proc.Reset(doc.Loader);

            // если нужно --> делаем Merge
            //doc.Merge();

            Log.exit();
            return doc;
        }
        /// <summary>
        /// getDoc(name)            - извлечение Документа name. Если еще не прочтен - из файла
        /// </summary>
        /// <param name="name">имя извлекаемого документа</param>
        /// <returns>Document</returns>
        /// <journal> 25.12.2013 отлажено
        /// 25.12.2013 - чтение из файла, формирование Range Body
        /// 28.12.13 - теперь doc.Sheet и doc.Wb храним в структуре Документа
        /// 5.1.14  - обработка шаблонов Документа
        /// 7.1.14  - отделяем пятку и помещаем в Body и Summary
        /// </journal>
        public static Document getDoc(string name)
        {
            Log.set("getDoc(" + name + ")");
            try
            {
                Document doc = Documents[name];
                if (!doc.isOpen)
                {
                    //---- загрузка Документа из файла
                    doc.Wb = FileOpenEvent.fileOpen(doc.FileName);
                    doc.Sheet = doc.Wb.Worksheets[doc.SheetN];
                    doc.splitBodySummary();
                    int newEOL = doc.Body.iEOL();
                    if (newEOL != doc.EOLinTOC)
                    {
                        Log.Warning("переопределил EOL(" + name + ")="
                            + newEOL + " было " + doc.EOLinTOC);
                        doc.EOLinTOC = newEOL;
                    }
                    if (!doc.stamp.Check(doc))
//                    if (!doc.stamp.Check())
                    {
                        new Log("Fatal Stamp chain");

                        //  НЕ НАПИСАНО!!!                      Stamp.trace(rng, 
                    }
                    doc.isOpen = true;
                }
                Log.exit();
                return doc;
            }
            catch
            {
                Log.FATAL("Документ \"" + name + "\" не открыт!!");
                // надо проверить, что Document name не существует
                // в случае, если существует, но не удалось прочитать - создать событие FATAL_ERR
                return null;    // нужно только при обработке Event File Open для неизвестного файла
            }
        }
        /// <summary>
        /// отделение основной части Документа (doc.Body) от пятки (doc.Summary)
        /// </summary>
        private void splitBodySummary()
        {      
            int fullEOL = Lib.MatchLib.EOL(Sheet);
            int _resLns = 0;
            switch (ResLines.Count)
            {
                case 0: break;
                case 1: _resLns = ResLines[0]; break;
                default: _resLns = (this.MadeStep == "Loaded") ? ResLines[0] : ResLines[1]; break;
            }
            int iEOL = (_resLns == 0) ? fullEOL : fullEOL - _resLns;
            int iEOC = Lib.MatchLib.EOC(Sheet);

            Body = FileOpenEvent.getRngValue(Sheet, 1, 1, iEOL, iEOC);
            if (_resLns > 0) Summary = FileOpenEvent.getRngValue(Sheet, iEOL + 1, 1, fullEOL, iEOC);
        }
        /// <summary>
        /// isDocChanged(name) - проверяет, что Документ name доступен и изменен
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <juornal> 10.12.2013
        /// 30.12.13 - проверка, что Документ не только существует, но изменился
        /// </juornal> 
        public bool isDocChanged(string name)
        {
            bool result = Documents.ContainsKey(name);
            if (!result)
            {
                new Log("FATAL -- вызван несуществующий Документ \"" + name + "\"");
                return false;
            }
            Document doc = Documents[name];
            return doc.isOpen && doc.isChanged;
        }
        /// <summary>
        /// "сброс" Документа приводит к тому, что его содержимое выбрасывается,
        /// остаются только заголовки колонок.
        /// </summary>
        /// <journal>9.1.2014</journal>
        public void Reset()
        {
            Log.set("Reset");
/* PK            try
            {
                Body.Rows["2:" + Body.Rows.Count].Delete();
//                Body.Range["A2", Body.Cells[Body.Rows.Count, 1]].Delete();
/////                saveDoc();
            }
            catch { Log.FATAL("Ошибка при удалении строк Body Документа \"" + name + "\""); }
            finally { Log.exit(); }
 PK */
        }
 /* PK       /// <summary>
        /// добавляет строку к Body Документа
        /// </summary>
        /// <journal>9.1.2014</journal>
        public Excel.Range AddRow()
        {
            Log.set("AddRow");
            try
            {
                Body.Range["A" + (int)(Body.Rows.Count + 1)].EntireRow.Insert();
//                Body.Rows[Body.Rows.Count].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
//                return Body.Rows[Body.Rows.Count];
                return Body;
            }
            catch
            { 
                Log.FATAL("Ошибка при добавлении строки Документа \"" + name + "\"");
                return null;
            }
            finally { Log.exit(); }
        }
 PK */
        /// <summary>
        /// сохраняет Документ, если он изменялся
        /// </summary>
        /// <param name="name"></param>
        public void saveDoc()
        {
            if (isOpen && isChanged) FileOpenEvent.fileSave(Wb);
        }
        public static void saveDoc(string name)
        {
            Document doc = Documents[name];
            if (doc.isDocChanged(name)) FileOpenEvent.fileSave(doc.Wb);
        }
        /// <summary>
        /// recognizeDoc(wb)        - распознавание Документа в Листе[1] wb
        /// </summary>
        /// <param name="wb"></param>
        /// <returns>имя распознанного документа или null, если Документ не распознан</returns>
        /// <journal> 14.12.2013
        /// 16.12.13 (ПХ) переписано распознавание с учетом if( is_wbSF(wb) )
        /// </journal>
        public static string recognizeDoc(Excel.Workbook wb)
        {
            Log.set("recognizeDoc(wb)");
            Excel.Worksheet wholeSheet = wb.Worksheets[1];
//            Excel.Range rng = wholeSheet.Range["1:" + Lib.MatchLib.EOL(wholeSheet).ToString()];
            object[,] toc = wholeSheet.UsedRange.get_Value();   //перенос данных из Excel в память
            int iEOL = toc.GetLength(0);
            int iEOC = toc.GetLength(1);
/* PK
            Mtrx[,] mtrx = new Mtrx[iEOL + 1, iEOC + 1];
            for (int i = 1; i <= iEOL; i++)
            {
                for (int j = 1; j <= iEOC; j++)
                {
                    mtrx[i, j] = new Mtrx(toc[i, j]);
                }
            }
*/
            Stamp stmpSF = Documents["SFDC"].stamp;
//--            bool is_wbSF = stmpSF.Check(mtrx);
//            bool is_wbSF = Stamp.Check(mtrx, stmpSF);
            // ищем подходящий документ в TOCmatch
            foreach (var doc in Documents)
            {
//--                if (is_wbSF && (doc.Value.FileName != Decl.F_SFDC)) continue;
                if (doc.Value.name == "SFDC" || doc.Value.name == "Process") continue;
//--                if (doc.Value.stamp.Check(mtrx))
//                if (Stamp.Check(rng, doc.Value.stamp))
                    {
                    Log.exit();
                    return doc.Value.name;
                }
            }       // конец цикла по документам
            return null;        // ничего не нашли
        }
        /// <summary>
        /// инициирует Fetch-структуру Документа для Запроса fetch_rqst.
        /// Если fetch_rqst не указан - для всех Запросов Документа.
        /// </summary>
        /// <param name="fetch_rqst"></param>
        /// <journal>11.1.2014 PKh
        /// 15.1.2014 - дописан FetchInit() - просмотр всех Fetch Документа</journal>
        public void FetchInit()
        {
            Log.set("FetchInit");
            try
            {
                for (int col = 1; col <= ptrn.iEOC(); col++)
                {
                    string ftch = ptrn.String(Decl.PTRN_FETCH, col);
                    string[] ar = ftch.Split('/');
                    if (ar.Length <= 2) continue;
                    Document doc = getDoc(ar[0]);
                    doc.FetchInit(ftch);
                }
            }
            catch { Log.FATAL("ошибка FetchInit() для Документа \"" + name + "\""); }
            finally { Log.exit(); }
        }
        public void FetchInit(string fetch_rqst)
        {
            Log.set("FetchInit");
            try
            {
                if (String.IsNullOrEmpty(fetch_rqst)) { FetchInit(); return; }
                string[] ar_rqst = fetch_rqst.Split('/');
                if (!Documents.ContainsKey(ar_rqst[0])) Log.FATAL("нет такого Документа");
                string strFetch = ar_rqst[0] + "/" + ar_rqst[1];
                if (docDic.ContainsKey(strFetch)) return; // уже инициирован -> return
                Document doc = getDoc(ar_rqst[0]);
                string[] cols = ar_rqst[1].Split(':');
                int key = Lib.MatchLib.ToInt(cols[0]);
                int val = Lib.MatchLib.ToInt(cols[1]);
                Dictionary<string, string> keyDic = new Dictionary<string, string>();
                docDic.Add(strFetch, keyDic);
                DateTime t0 = DateTime.Now;
                for (int i = 1; i <= doc.Body.iEOL(); i++)
                {
                    string s1 = doc.Body.String(i, key);
                    if (s1 != "") keyDic.Add(s1, doc.Body.String(i, val));
                }
                DateTime t1 = DateTime.Now;
                new Log("-> "+(t1-t0));
            }
            catch { Log.FATAL("ошибка запроса \"" + fetch_rqst + "\" для Документа \"" + name + "\""); }
            finally { Log.exit(); }
        }

        /// <summary>
        /// Класс Stamp, описывающий все штампы документа
        /// </summary>
        /// <journal> дек 2013
        /// 12.1.2014 - работа с матрицей в памяти, а не с Range в Excel
        /// </journal>
//        private class Stamp : Document
        private class Stamp
        {
            public List<OneStamp> stamps = new List<OneStamp>();
            /// <summary>
            /// конструируем цепочку Штампов по строкам TOC от i0 до i1
            /// </summary>
            /// <param name="i0"></param>
            /// <param name="i1"></param>
            public Stamp(int i0, int i1)
            {
                Document doc_toc = getDoc(TOC);
                if (doc_toc.Body.String(i0, Decl.DOC_STMPTYPE) != "N")
                {
                    for (int i = i0; i <= i1; i++) stamps.Add(new OneStamp(doc_toc, i));
                }
            }
            /// <summary>
            /// Check(Документ) - проверка, что штамп в Документе соответствует цепочке Штампов в TOCmatch
            /// </summary>
            /// <param name="doc">проверяемый Документ</param>
            /// <returns>true, если результат проверки положительный, иначе false</returns>
            /// <journal> 12.12.13
            /// 16.12.13 (ПХ) перенес в класс Stamp и переписал
            /// 13.1.2014 - переписано
            /// </journal>
//            public static bool Check(Excel.Range rng, Stamp stmp)
            public bool Check(Document doc)
            {
                Mtr range = (doc.FileName == Decl.F_SFDC) ? doc.Summary : doc.Body;
                foreach (OneStamp st in stamps) if (!st.Check(range)) return false;
                return true;
            }

            //public bool Check()
            //{
            //    Mtr range = (FileName == Decl.F_SFDC) ? Summary : Body;
            //    foreach (OneStamp st in stamps) if (!st.Check(range)) return false;
            //    return true;
            //}
            /* PK
                        /// <summary>
                        /// trace(Stamp)    - вывод в Log-файл данных по Штампам Документа
                        /// </summary>
                        /// <param name="st"></param>
                        /// <journal> 26.12.13 -- не дописано -- нужно rnd не только doc.Body, но для SF doc.Summary
                        /// </journal>
                        public void trace(Document doc)
                        {
                            Log.set("Stamp.trace(" + doc.name + ")");
                            Excel.Range rng = (doc.FileName == Decl.F_SFDC) ? doc.Summary : doc.Body;
                            foreach (OneStamp st in doc.stamp.stamps)
                                if (OneStamp.Check(rng, st))
                                {
                                    new Log("\t=OK=>" + st.ToString());
                                }
                                else
                                {
                                    new Log("\t=!!=>" + st.ToString() + "\tFATAL!");
                                }
                            Log.FATAL("Документ соответствует Штампам");
                        }
            PK */
        }

        /// <summary>
        /// Класс, описывающий штамп документа (с вариантами позиций, заданными в одной стрке TOCmatch)
        /// </summary>
//        public class OneStamp : Document
        public class OneStamp
        {
            private string signature;   // проверяемый текст Штампа - сигнатура
            private string typeStamp;     // '=' - точное соответствие сигнатуры; 'I' - "текст включает.."
            private List<int[]> stampPosition = new List<int[]>();   // альтернативные позиции сигнатур Штампов
//            private bool _isSF;

            /// <summary>
            /// Конструктор OneStanp(doc_toc, int rowNumber)
            /// </summary>
            /// <param name="doc_toc">  таблица TOCmatch</param>
            /// <param name="rowNumber">одна строка штампа (т.е. сигнатура и позиции)</param>
            /// <example>
            /// примеры: {[1, "1, 6"]} --> [1,1] или [1,6]
            ///  .. {["4,1", "2,3"]} --> [4,2]/[4,3]/[1,2]/[1,3]
            /// </example>
            /// <journal> 12.12.2013 (AP)
            /// 16.12.13 (ПХ) добавлен параметр isSF - добавляется в структуру Штампа
            /// 12.1.14 - работаем с TOCmatch с памяти -- Matrix
            /// </journal>
            public OneStamp(Document doc, int rowNumber)
            {
                signature = doc.Body.String(rowNumber, Decl.DOC_STMPTXT);
                typeStamp = doc.Body.String(rowNumber, Decl.DOC_STMPTYPE);
  
                List<int> rw = intListFrCell(doc, rowNumber, Decl.DOC_STMPROW);
                List<int> col = intListFrCell(doc, rowNumber, Decl.DOC_STMPCOL);
                // декартово произведение множеств rw и col
                rw.ForEach(r => col.ForEach(c => stampPosition.Add(new int[] { r, c })));
            }
            /// <summary>
            /// Check(rng, stmp)        - проверка сигнатуры Штампа stmp в rng для его всех допустимых позиций
            /// </summary>
            /// <param name="rng"></param>
            /// <param name="stmp"></param>
            /// <returns>bool: true если проверка Штампа дает совпадение сигнатуры</returns>
            /// <journal> 12.12.2013
            /// 25.12.13 (ПХ) ToString вместо Value2 для проверяемой ячейки
            /// 13.1.14 - работа с матрицами
            /// </journal>
            public bool Check(Mtr mtr)
            {
 //               int shiftToEol = (Wb.Name == Decl.F_SFDC) ? rng.GetLength(0) - 6 : 0;
                string sig = signature.ToLower();
                foreach (var pos in stampPosition) {
                    var x = mtr.String(pos[0], pos[1]);
                    string strToCheck = x.ToLower();
                    if (typeStamp == "=") {
                        if (strToCheck == sig) return true;
                    } else {
                        if (strToCheck.Contains(sig)) return true;
                    }
                }
                return false;
            }
                
//                string sig = stmp.signature.ToLower();
//                foreach (var pos in stmp.stampPosition)
//                {
////                    var x = rng.Cells[pos[0] + shiftToEol, pos[1]].Value2.ToString();
//                    if (x == null) continue;
//                    string strToCheck = x.ToLower();

//                    if (stmp.typeStamp == '=')
//                    {
//                        if (strToCheck == sig) return true;
//                    }
//                    else
//                    {
//                        if (strToCheck.Contains(sig)) return true;
//                    }
//                }

            private List<int> intListFrCell(Document doc, int row, int col)
            {
                return Lib.MatchLib.ToIntList(doc.Body.String(row, col), ',');
            }
        }   // конец класса OneStamp
    }    // конец класса Document
}