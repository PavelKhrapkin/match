/*-----------------------------------------------------------------------
 * Document -- класс Документов проекта match 3.1
 * 
 *  31.01.2015  П.Храпкин, А.Пасс
 *  
 * -------------------------------------------
 * Document(name)       - КОНСТРУКТОР возвращает объект Документ с именем name
 * loadDoc(name, wb)    - загружает Документ name или его обновления из файла wb, запускает Handler Документа
 * getDoc(name)         - возвращает Документ с именем name; при необходимости - открывает его
 * NewSheet(name)       - созданние нового листа с заголовком для Документа name
 * isDocChanged(name)   - проверяет, что Документ name открыт и изменился позле загрузки
 * saveDoc()            - сохраняет Документ, если он изменялся
 * recognizeDoc(wb)     - распознает первый лист файла wb по таблице Штампов
 * 
 * внутренний класс Stamp предназначен для заполнения списков Штампов
 * каждый Штамп содержит сигнатуру, то есть проверяемый текст, и пару координат - его положений
 * Stamp(Range rng)     - разбирает rng, помещая из таблицы TOCmatch Штамп в List Штампов в Документе
 * Check(rng,stampList) - проверка Штампов stampList в Range rng 
 */
using System;
using System.Data;
//using Box = System.Windows.Forms.MessageBox;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using FileOp = match.MyFile.FileOpenEvent;
using Decl = match.Declaration.Declaration;
//using Mtrx = match.Matrix.Matrix;
using Mtr = match.Matrix.Matr;  //--
using Lib = match.Lib;
using CS = match.Lib.CS;
using Log = match.Lib.Log;
using Proc = match.Process.Process;
using System.Xml.Serialization;

namespace match.Document
{
    /// <summary>
    /// класс Document содержит таблицу параметров всех Документов, известных приложению match
    /// </summary>
    [XmlRoot]
    public class Document
    {
        private static Dictionary<string, Document> Documents = new Dictionary<string, Document>();   //коллекция Документов

        public string name;
        private bool isOpen = false;
        public bool isChanged = false;
#if WB_PRIVATE
        //private Excel.Workbook wb;
        //public Excel.Workbook Wb { get; private set; }
#else
        [XmlIgnore]
        public Excel.Workbook Wb;
#endif
        private string FileName;
        private string SheetN;
        [XmlIgnore]
        public Excel.Worksheet Sheet;
        private string MadeStep;
        private DateTime MadeTime;
        private int periodDays;     // периодичность работы с Документом в днях
        private double chkSum;      // контрольная сумма
        private long colorTab;      // цвет Tab листа
        private int EOLinTOC;
        private List<int> ResLines; //число строк в пятке -- возможны альтернативные значения
        private Stamp stamp;        //каждый документ ссылается на цепочку сигнатур или Штамп
        private DateTime creationDate;  // дата создания Документа 
        private string Loader;
        private string LastUpdateFromFile;
//        private bool isPartialLoadAllowed;
        public int MyCol;           // количесто колонок, добавляемых слева в Документ в loadDoc
        public int usedColumns;     // общее кол-во использованных колонок в Body Документа
        [XmlIgnore]
        public Mtr ptrn;            //шаблон - заголовок таблицы
        [XmlIgnore]
        public Mtr Body;
        public DataTable dt;
        [XmlIgnore]
        public Mtr Summary;
        [XmlIgnore]
        public Dictionary<string, Dictionary<string, string>> docDic = new Dictionary<string, Dictionary<string, string>>();

        private const string TOC = "TOCmatch";
        private const int TOC_DIRDBS_COL = 10;  //в первой строке в колонке TOC_DIRDBS_COL записан путь к dirDBs
        private const int TOC_LINE = 4;         //строка номер TOL_LINE таблицы ТОС отностися к самому этому документу.

        static Document()
        {
            Log.set("статический конструктор Документов");
            Document doc = null;
            Excel.Workbook db_match = FileOp.fileOpen(Decl.F_MATCH);
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
                    doc.colorTab = FileOp.cellColorIndex(tocSheet, i, Decl.DOC_SHEET);
                    doc.FileName = mtr.String(i, Decl.DOC_FILE);
                    doc.SheetN   = mtr.String(i, Decl.DOC_SHEET);
                    doc.creationDate = Lib.MatchLib.getDateTime(mtr.get(i, Decl.DOC_CREATED));
                    string ptrnName = mtr.String(i, Decl.DOC_PATTERN);
                    try {
                        if (ptrnName != "") doc.ptrn = new Mtr(hdrSht.Range[ptrnName].get_Value());
                    } catch {
                        if (ptrnName == "WP_Prototype") continue;
                    } // затычка специально для WP_prototype 23.1.14
                    doc.Loader = mtr.String(i, Decl.DOC_LOADER);
                    int j;
                    for (j = i + 1; j <= iEOL && mtr.String(j, Decl.DOC_NAME) == ""; j++) ;
                    doc.stamp = new Stamp(i, j - 1);
                } //if docName !=""
            } // for по строкам TOC

 //                   try { docAcc.creationDate = Lib.MatchLib.getDateTime(Double.Parse(rw.Range[Decl.DOC_CREATED].Text)); }
 //                   catch { docAcc.creationDate = new DateTime(0); }
 //                   try { docAcc.ptrn = hdrSht.get_Range((string)rw.Range[Decl.DOC_PATTERN].Text); } catch { docAcc.ptrn = null; }
 //                   try { docAcc.SummaryPtrn = hdrSht.get_Range((string)rw.Range[Decl.DOC_SUMMARY_PATTERN].Text); } catch { docAcc.SummaryPtrn = null; }
 //                   docAcc.Loader = rw.Range[Decl.DOC_LOADER].Text;
 //                   // флаг, разрешающий частичное обновление Документа пока прописан хардкодом
 //                   switch (docName)
 //                   {
 //                       case "Платежи":
 //                       case "Договоры": docAcc.isPartialLoadAllowed = true;
 //                           break;
 //                       default: docAcc.isPartialLoadAllowed = false;
 //                           break;
 //                   }
 //               }
 //           }

////////////            //-----------------------------------------------------------------
////////////            // из коллекции Documents переносим произошедшие изменения в файл
////////////            //            if (docAcc.Body.Range["A" + TOC_DIRDBS_COL].Value2 != Decl.dirDBs)
////////////            {
////////////                //    Box.Show("Файл '" + F_MATCH + "' загружен из необычного места!");
////////////                //    // переустановка match -- будем делать потом
////////////                docAcc.isChanged = true;
////////////            }
//////////////            docAcc.EOLinTOC = iEOL;
//////////////PK            docAcc.Body.Range["C4"].Value2 = iEOL.ToString();
////////////            docAcc.isChanged = true;   // TOCmatch сохраняем всегда. Возможно, это времянка
//////////////            docAcc.isOpen = true;
////////////            docAcc.saveDoc();
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
/*            if (docAcc.isPartialLoadAllowed)
            {
                // Дальше распознавание частичное или полное обновление прежнего Документа.
                // Здесь только если частичное, то есть потом будет выполняться Merge
                // это еще не написано!!
            }
 */
            doc.LastUpdateFromFile = wb.Name;
            string oldRepName = "Old_" + doc.SheetN;
            try
            {
                wb.Worksheets[1].Name = "TMP";
                wb.Worksheets[1].Move(doc.Sheet);
                 // если Old_ уже есть, но еще не обработан - уничтожаем прежний "частичный" отчет
                if (FileOp.sheetExists(doc.Wb, oldRepName))
                {
                    FileOp.DisplayAlert(false);
                    doc.Wb.Worksheets[doc.name].Delete();
                    FileOp.DisplayAlert(true);
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
            doc.FetchInit();

            ////// если есть --> запускаем Handler
            ////if (docAcc.Loader != null) Proc.Reset(docAcc.Loader);
            ////// если нужно --> делаем Merge name с oldRepName
            ////if (FileOp.sheetExists(docAcc.Wb, oldRepName))
            ////{
            ////    // еще не написано!!
            ////    // NB: в таблице Процессов есть Шаг MergeReps
            ////}
            Log.exit();
            return doc;
        }
        /// <summary>
        /// getDoc(name)            - извлечение Документа name. Если еще не прочтен - из файла
        /// </summary>
        /// <param name="name">имя извлекаемого документа</param>
        /// <returns>Document или null? если Документ не найден</returns>
        /// <journal> 25.12.2013 отлажено
        /// 25.12.2013 - чтение из файла, формирование Range Body
        /// 28.12.13 - теперь docAcc.Sheet и docAcc.Wb храним в структуре Документа
        /// 5.1.14  - обработка шаблонов Документа
        /// 7.1.14  - отделяем пятку и помещаем в Body и Summary
        /// 5.4.14  - инициализируем docDic, то есть подготавливаем набор данных для Fetch
        /// 31.1.15 - вначале проверка, есть ли Документ name 
        /// </journal>
        public static Document getDoc(string name)
        {
            Log.set("getDoc(" + name + ")");
            try
            {
                if (!Documents.ContainsKey(name)) return null;  //специально для UniTest
                Document doc = Documents[name];
                if (!doc.isOpen)
                {
                    //---- загрузка Документа из файла
                    doc.Wb = FileOp.fileOpen(doc.FileName);
                    doc.Sheet = doc.Wb.Worksheets[doc.SheetN];
                    doc.splitBodySummary();
                    int newEOL = doc.Body.iEOL();
                    if (newEOL != doc.EOLinTOC)
                    {
                        Log.Warning("переопределил EOL(" + name + ")="
                            + newEOL + " было " + doc.EOLinTOC);
                        doc.EOLinTOC = newEOL;
                    }
                    Mtr rng = (doc.FileName == Decl.F_SFDC)? doc.Summary: doc.Body;
                    if (!doc.stamp.Check(rng)) Log.FATAL(doc.stamp.Trace(doc));
                    doc.isOpen = true;
                }
                Log.exit();
                return doc;
            }
            catch
            {
                string msg = (Documents.ContainsKey(name)) ? "не существует" : " не удалось открыть";
                Log.FATAL("Документ \"" + name + "\" " + msg);
                return null;    // нужно только при обработке Event File Open для неизвестного файла
            }
        }
        /// <summary>
        /// NewSheet(name)  - созданние нового листа с заголовком для Документа name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>вновь созданный Документ name</returns>
        /// <journal>6.4.2014
        /// 25.12.14 дописано
        /// 13.01.15 bug fix: присваиваем новый docAcc.Sheet
        /// </journal>
        public static Document NewSheet(string name)
        {
            Log.set("NewSheet(" + name + ")");
            Document doc = getDoc(name);
            if (doc == null) Log.FATAL("не создан новый лист в NewSheet(" + name + ")");
            try
            {
                Excel.Workbook wb = doc.Wb;
                //-- заменяем лист Sheet на пустой
                 wb.Application.DisplayAlerts = false;
                    doc.Sheet.Delete();
                    wb.Application.Sheets.Add(After: wb.Application.Sheets.Item[2]);
                    Excel.Worksheet Sh = wb.Application.Sheets[3];
                    doc.Sheet = Sh;              
                    Sh.Name = name;
                    wb.Application.Sheets[3].Tab.ColorIndex = doc.colorTab;
                wb.Application.DisplayAlerts = true;
                //-- в новый Body переносим строку - заголовок из ptrn = Header
                object[,] hh = new object[1, doc.ptrn.iEOC()];
                for (int i = 1; i <= doc.ptrn.iEOC(); i++) { hh[0,i-1] = doc.ptrn.String(1,i);}
                doc.Body = new Mtr(hh);
                //-- записываем в таблицу Documents данные по новому Документу name
                doc.isChanged = true;
                doc.creationDate = DateTime.Now;
                doc.MadeStep = "NewSheet"; doc.MadeTime = doc.creationDate;
                doc.EOLinTOC = 1;
            }
            catch(Exception er) { Log.FATAL("ошибка NewSheet(" + name + ") " + er); }
            finally { Log.exit(); }
            return doc;
        }
        /// <summary>
        /// отделение основной части Документа (docAcc.Body) от пятки (docAcc.Summary) и их чтение из Excel
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

            Body = FileOp.getRngValue(Sheet, 1, 1, iEOL, iEOC);
            if (_resLns > 0) Summary = FileOp.getRngValue(Sheet, iEOL + 1, 1, fullEOL, iEOC);
        }
        /// <summary>
        /// isDocChanged(name) - проверяет, что Документ name доступен и изменен
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <juornal> 10.12.2013
        /// 30.12.13 - проверка, что Документ не только существует, но изменился
        /// 17.01.15 - проверка не изменилась ли контрольной сумма this?
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
            double newChkSum = Lib.CS.CheckSum(doc);
            if (newChkSum != doc.chkSum) { doc.chkSum = newChkSum; doc.isChanged = true; }
            return doc.isOpen && doc.isChanged;
        }

        /// <summary>
        /// saveDoc() - сохраняет Документ, если он изменялся
        /// </summary>
        /// <journal>14.1.2015 PKh
        /// 17.01.2015 - проверяем контрольную сумму Документа перед сохранением
        /// </journal>
        /// <param name="name"></param>
        public void saveDoc()
        {
            Log.set("saveDoc(" + name + ")");
            if (isDocChanged(this.name))
            {        
                int rows = Body.iEOL();
                int cols = Body.iEOC();
                var startCell = (Excel.Range)this.Sheet.Cells[1, 1];
                var endCell = (Excel.Range)this.Sheet.Cells[rows, cols];
                Excel.Range writeRange = Sheet.Range[startCell, endCell];
 // ---- реализуем     Excel.Range r = (Excel.Range)this.Body;
 // ---- для этого перекладываем поэлементно из Body в data
                var data = new object[rows, cols];
                for (var row = 1; row <= rows; row++)
                {
                    for (var column = 1; column <= cols; column++)
                    {
                        data[row - 1, column - 1] = Body.get(row - 1, column - 1);
                    }
                }
                writeRange.Value2 = data;
                if (name == TOC || name == Decl.PROCESS) data[1, 1] = DateTime.Now;
// --- устанавливаем ширину колонки листа по значениям в строке Шаблона Width
                for (int col = 1; col <= cols; col++)
                {
                    string wdth = ptrn.String(3, col);
                    string[] ar = wdth.Split('/');
                    float W;
                    if (!float.TryParse(ar[0], out W)) Log.FATAL("ошибка в строке Width шаблона = \""
                        + wdth + "\" при обработке Документа " + this.name);
                    writeRange.Columns[col].ColumnWidth = W;
                }
// -!!- еще надо записать docAcc.Summary, но это реализуем позже

                FileOp.fileSave(this.Wb);
// -??- надо убедиться, что в TOCmatch правильно записывается изменение состояния Документа
                if (isDocChanged(TOC)) Documents[TOC].saveDoc(); //рекурсия!!
            }
            Log.exit();
        }
        private static void colCpy(Mtr mtr, int rwMtr, Excel.Range rng, int rwRng)
        {
            int cols = mtr.iEOC();
            for (int col = 1; col <= cols; col++) rng.Cells[rwRng, col] = mtr.get(rwMtr, col);
        }
        public static string recognizeDoc(Excel.Workbook wb)
        {
            Log.set("recognizeDoc(wb)");
            Mtr wbMtr = FileOp.getSheetValue(wb.Worksheets[1]);
            // вначале проверим где у wb Штамп - в теле или в пятке? Штамп в пятке бывает только у SF
            // отделим от wbMtr область пятки SF -- переложим SFresLines строк wbMtr в wdSFsummary
            int iEOL = wbMtr.iEOL();
            int iEOC = wbMtr.iEOC();
            object[,] tmp = new object [Decl.SFresLines + 1, iEOC + 1];
            for (int rw = 1; rw <= Decl.SFresLines; rw++)
                for (int col = 1; col <= iEOC; col++)
                    tmp[rw, col] = wbMtr.get(iEOL - Decl.SFresLines + rw - 1, col);
            Mtr wbSFsummary = new Mtr(tmp);

            Mtr rng = (Documents["SFDC"].stamp.Check(wbSFsummary))? wbSFsummary: wbMtr;

            try 
            {
                foreach (var doc in Documents)  // ищем подходящий документ по Штампам
                    if (doc.Value.stamp.Check(rng)) return doc.Value.name;
                return null;                    // если ничего не нашли -> вовращаем null
            }
            finally { Log.exit(); }                  
        }
        /// <summary>
        /// ! не дописано!
        /// AddLine(dbLine, NewItemsDoc, ext) - добавляет в NewItemsDoc запись, обрабатывая строку dbLine
        /// </summary>
        /// <param name="dbLine">object dbLine - обрабатываемая строка Документа</param>
        /// <param name="NewItemsDoc">Docs NewItemsDoc - Документ для переноса в SF</param>
        /// <param name="ext">ext - опциональный "внешний" параметр или массив параметров</param>
        /// <journal> 31.01.2015</journal>
        public void AddLine(DataRow dbLine, Document NewItemsDoc, object Ext)
        {
            object[] rw = new object[100];
            DataTable ptrn = NewItemsDoc.ptrn.DaTab();
            int i = 1;
            DataRow pColmn = ptrn.Rows[3];
            DataRow pAdapt = ptrn.Rows[4];
            DataRow pFetch = ptrn.Rows[5];
            foreach (DataColumn col in ptrn.Columns)
            {
                var x = dbLine[ Convert.ToInt16(pColmn[col]) ].ToString();
                string ftchRqst = pFetch[col].ToString();
                if (!String.IsNullOrEmpty(ftchRqst) ) x = Fetch(ftchRqst, x);
                string adpt = pAdapt[col].ToString();
                if (String.IsNullOrEmpty(adpt)) rw[i] = x;
   //             else rw[i] = AdaptEngine()
                i++;
            }
            NewItemsDoc.dt.Rows.Add(rw);
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


                    string[] br = ftch.Split((char)'/');



                    doc.FetchInit(ftch);
                }
            }
            catch { Log.FATAL("ошибка FetchInit() для Документа \"" + name + "\""); }
            finally { Log.exit(); }
        }
 
        /// <param name="fetch_rqst"></param>
        /// <example>FetchInit("SFacc/2:3")</example>
        public void FetchInit(string fetch_rqst)
        {
            Log.set("FetchInit(fetch_rqst)");
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
                    if (s1 != "")try { keyDic.Add(s1, doc.Body.String(i, val)); }
                        catch
                        {
                            Log.Warning("Запрос \"" + fetch_rqst + " Строка " + i
                                + " неуникальное значение \"" + s1 + "\" в ключевом поле запроса!");
                        }
                }
                DateTime t1 = DateTime.Now;
                new Log("-> "+(t1-t0));
            }
            catch { Log.FATAL("ошибка запроса \"" + fetch_rqst + "\" для Документа \"" + name + "\""); }
            finally { Log.exit(); }
        }

        /// <summary>
        /// Fetch(fetch_rqst, x) -- извлекает значение по строкам х и ftch_rqst
        /// </summary>
        /// <example>Fetch("SFacc/2:3/0", "ООО «ОРБИТА СПб»") </example>
        /// <example>Fetch("SF/2:3/0", "ООО «ОРБИТА СПб»") </example>
        /// <param name="fetch_rqst"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        /// <journal>5.4.2014
        /// 31/01/2015 - object fetch_reqst, x; if(fetch_reqst == "" || x =="") return x;
        /// </journal>
        public string Fetch(object fetch_rqst, object x)
        {
            Log.set("Fetch");
            string result = null;
            try
            {
                string sfr = fetch_rqst.ToString().Trim(), sx = x.ToString().Trim();
                if (sfr == null || sx == null) Log.FATAL("параметр null"); 
                result = sx;
                if ( sfr != "" )
                {
                    string[] ar_rqst = sfr.Split('/');
                    Document doc = getDoc(ar_rqst[0]);
                    Dictionary<string, string> Dic = doc.docDic[ar_rqst[0] + "/" + ar_rqst[1]];
                    result = Dic[sx];
                }
            }
            catch { Log.FATAL("ошибка Fetch( \"" + fetch_rqst + "\", \"" + x + "\")" ); }
            finally { Log.exit(); }
            return result;
        }

        /// <summary>
        /// Класс Stamp, описывающий все штампы документа
        /// </summary>
        /// <journal> дек 2013
        /// 12.1.2014 - работа с матрицей в памяти, а не с Range в Excel
        /// </journal>
        private class Stamp
        {
            public List<OneStamp> stamps = new List<OneStamp>();
            /// <summary>
            /// конструируем цепочку Штампов по строкам TOC от i0 до i1
            /// </summary>
            /// <param name="i0"></param>
            /// <param name="i1"></param>
            /// <journal>
            /// 18.1.2014 (ПХ) в класс Штамп и в конструктор добавлен _parentDoc - Документ Штампа
            /// </journal>
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
            /// 
            /// Штамп.Check(Mtr) - проверяем, что данные в Mtr содержат сигнатуры Штампа на нужных местах
            /// </summary>
            /// <param name="docAcc">проверяемый Документ</param>
            /// <returns>true, если результат проверки положительный, иначе false</returns>
            /// <journal> 12.12.13
            /// 16.12.13 (ПХ) перенес в класс Stamp и переписал
            /// 13.1.2014 - переписано
            /// 18.1.14 (ПХ) - переписано еще раз: проверяем mtr
            /// </journal>
            public bool Check(Mtr mtr)
            {             
                if (mtr == null) return false;
                foreach (OneStamp st in stamps) if (!st.Check(mtr)) return false;
                return true;
            }
            /// <summary>
            /// Trace(Stamp)    - вывод в Log-файл данных по Штампам Документа
            /// </summary>
            /// <param name="st"></param>
            /// <journal> 26.12.13 -- не дописано -- нужно rnd не только docAcc.Body, но для SF docAcc.Summary
            /// 18.1.14 (ПХ) отладка с Matrix
            /// </journal>
            public string Trace(Document doc)
            {
                Mtr rng = (doc.FileName == Decl.F_SFDC) ? doc.Summary : doc.Body;
                string msg = (string)((rng == doc.Summary) ? "Пятка" : "Body");
                msg += "Документ не соответствует Штампам";
                foreach (OneStamp st in doc.stamp.stamps)
                    traceSub(st.Check(rng) ? "OK" : "!!", st);
                return msg;
            }
            static void traceSub(string msg, OneStamp st)
            {
                new Log("\t=" + msg + "=> " + st.get("type") + " " + st.get("sig") + " " + st.get());
            }
        }
        /// <summary>
        /// Класс, описывающий штамп документа (с вариантами позиций, заданными в одной стрке TOCmatch)
        /// </summary>
        public class OneStamp
        {
            private string signature;   // проверяемый текст Штампа - сигнатура
            private string typeStamp;   // '=' - точное соответствие сигнатуры; 'I' - "текст включает.."
            private List<int[]> stampPosition = new List<int[]>();   // альтернативные позиции сигнатур Штампов
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
  
                List<int> rw  = intListFrCell(doc, rowNumber, Decl.DOC_STMPROW);
                List<int> col = intListFrCell(doc, rowNumber, Decl.DOC_STMPCOL);
                // декартово произведение множеств rw и col
                rw.ForEach(r => col.ForEach(c => stampPosition.Add(new int[] { r, c })));
            }
            /// <summary>
            /// используется для внешнего доступа к private полям Штампа, в т.ч. для Log и Trace
            /// </summary>
            /// <param name="str">что извлекаем: "signature" или "type" или "position"</param>
            /// <returns>string</returns>
            /// <journal> 18.1.2014 (ПХ)</journal>
            public string get(string str = "position")
            {
                string v;
                switch (str.ToLower()[0])
                {
                    case 's': v = signature; break;
                    case 't': v = typeStamp; break;
                    default:
                        {
                            v = "{";
                            foreach (int[] pos in stampPosition) v += "[" + pos[0] + "," + pos[1] + "]";
                            v += "}"; break;
                        }
                }
                return v;
            }
            private List<int> intListFrCell(Document doc, int row, int col)
            {
                return Lib.MatchLib.ToIntList(doc.Body.String(row, col), ',');
            }
            /// <summary>
            /// Stamp.Check(mtr) - проверка mtr на соответствие Штампу в stmp. ОК->true, несовпадение->false
            /// </summary>
            /// <param name="mtr"></param>
            /// <returns>bool: true если проверка Штампа дает совпадение сигнатуры</returns>
            /// <journal> 12.12.2013
            /// 25.12.13 (ПХ) ToString вместо Value2 для проверяемой ячейки
            /// 13.1.14 - работа с матрицами
            /// 18.1.14 - (ПХ) рефакторинг. Теперь сверяем strToCheck в mtr и SigInStamp в Штампе
            /// </journal>
            public bool Check(Mtr mtr)
            {
                string sigInStamp = signature.ToLower();
                foreach (var pos in stampPosition) {
                    string strToCheck = mtr.String(pos[0], pos[1]).ToLower();
                    if (typeStamp == "=") {
                        if (strToCheck == sigInStamp) return true;
                    } else {
                        if (strToCheck.Contains(sigInStamp)) return true;
                    }
                }
                return false;
            }
        }   // конец класса OneStamp
    }    // конец класса Document

    /// <summary>
    /// формирование списков строк List<string> по Организациям, Контактам, Проектам и т.п. 
    /// </summary>
    /// <journal>22.1.2015</journal>
    public class Lst
    {
        public static List<string> Accounts = new List<string>();  //коллекция Организаций в SF
        public static List<string> Acc1Cs = new List<string>();    //коллекция Организаций в 1С
        public static List<string> Contacts = new List<string>();  //коллекция Контактов
        public static List<string> Opps = new List<string>();      //коллекция Проектов
        public static List<string> Pays = new List<string>();      //коллекция Платежей
        public static List<string> Contracts = new List<string>(); //коллекция Договоров  

        public enum Entity { Accounts, Acc1Cs, Contacts, Opps, Pays, Contracts };
        static Lst() {} //конструктор не нужен -- листы статические

        public static void Init(Entity ent)
        {
            switch (ent)
            {
                case Entity.Accounts:   Accounts =  docToList("SFacc", 0); break;
                case Entity.Acc1Cs:     Acc1Cs =    docToList("Список клиентов 1C", 3); break;
                case Entity.Contacts:   Contacts =  docToList("SFcont", 4); break;
                case Entity.Opps:       Accounts =  docToList("SFopp", 1); break;
                case Entity.Pays:       Pays     =  docToList("Платежи", 5); break;
                case Entity.Contracts:  Contracts = docToList("Договоры", 10); break;
            }
        }
        private static List<string> docToList(string name, int ind)
        {
            List<string> _Lst = new List<string>();
            Document doc = Document.getDoc(name);
            doc.dt = doc.Body.DaTab();
            foreach (DataRow Rw in doc.dt.Rows)
                _Lst.Add(Rw.Field<string>(ind));
            return _Lst;
        }


    }    // конец класса Lst
}