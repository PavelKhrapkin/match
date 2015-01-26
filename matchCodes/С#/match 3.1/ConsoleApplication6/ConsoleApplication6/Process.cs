/*-----------------------------------------------------------------------
 * Process -- класс Процессов проекта match 3.1
 * 
 *  17.01.2015  П.Храпкин, А.Пасс
 *  
 * -------------------------------------------
 * Process()            - КОНСТРУКТОР заполняет каталог процессов
 */
using System;
using System.Data;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using FileOp = match.MyFile.FileOpenEvent;
using Decl = match.Declaration.Declaration;
using Docs = match.Document.Document;
using match.Handler;
using match.Lib;
using Log = match.Lib.Log;

namespace match.Process
{
    /// <summary>
    /// класс Process (Процесс) содержит таблицу всех известных системе Процессов
    /// </summary>
    /// <journal> 29.12.2013
    /// 30.12.2013 - перенесено в C# из match 2.2 модуля ProcessEngin.bas
    /// 31.12.2013 -- основная реализация алгоритмов ProccessEngine
    /// 17.1.2015 - переписано без Data Table
    /// </journal>
    class Process
    {
        /// <summary>
        /// Processes - коллекция Процессов
        /// </summary>
        private static Dictionary<string, Process> Processes = new Dictionary<string, Process>();

        private string name;                            // имя Процесса
        private List<Step> steps = new List<Step>();    // Шаги, составляющие Процесс
        public List<int> results = new List<int>();     // результаты Процесса
        private int[] _rng;                             //номера строк <*>ProcStart и <*>ProcEnd Процесса
        private int line_procStart;

        /// <summary>
        /// Статический конструктор Процесса запускается системой перед первым обращением к классу.
        /// Он считывает данные в коллекцию Processes из листа Process файла match.xlsm
        /// </summary>
        /// <journal> 1.1.2014
        /// 17.1.2015 - переписано без Data Table
        /// </journal>
        static Process()
        {
            Log.set("статический конструктор Процессов");
            try {
                Process proc = null;
                Docs proc_doc = Docs.getDoc(Decl.PROCESS);
                for (int line_process = 0; line_process < proc_doc.Body.iEOL(); line_process++)
                {
                    int cols = proc_doc.Body.iEOC();
                    object[] rw = new object[cols];
                    for (int col=0; col < cols; col++) rw[col] = proc_doc.Body.get(line_process+1, col+1);
                    string cell = rw[Decl.STEP_NAME] as string;
                    if (String.IsNullOrEmpty(cell) || !String.IsNullOrEmpty(rw[Decl.STEP_COMMENT] as string)) continue;
                    switch (cell)
                    {
                        case "<*>ProcStart":
                            proc = new Process();
                            proc.line_procStart = line_process;
                            proc.name = rw[Decl.PROC_NAME] as string;
                            break;
                        case "<*>ProcEnd":
                            proc._rng = new int[] {proc.line_procStart, line_process};
                            Processes.Add(proc.name, proc);
                            break;
                        default:
                            if (proc == null || proc.name == null) continue;    //пустой Шаг
                            proc.steps.Add(new Step(rw));
                            break;
                    }
                }
                
                
                /* PK
                foreach (Excel.Range rw in proc_doc.Body.Rows)
                {
                    line_process++;
                    var cell = rw.Range[Decl.STEP_NAME].Text;
                    if (String.IsNullOrEmpty(cell) || !String.IsNullOrEmpty(rw.Range[Decl.STEP_COMMENT].Text)) continue;
                    switch (cell as string)
                    {
                        case "<*>ProcStart":
                            proc = new Process();
                            proc.line_procStart = line_process;
                            proc.name = rw.Range[Decl.PROC_NAME].Text;
                            break;
                        case "<*>ProcEnd":
                            proc._rng = proc_doc.Body.Range[proc.line_procStart + ":" + line_process];
                            Processes.Add(proc.name, proc);
#if MATCH_TRACE
                            new Log("\t- " + proc.name);
#endif
                            break;
                        default:
                            if (proc.name == null) continue;    //пустой Шаг
                            proc.steps.Add(new Step(rw));
                            break;
                    }
                }
 PK */
                new Log("после работы статического конструктора Процессов");
            }
            catch { Log.FATAL("Не удалось инициализировать таблицу Процессов."); }
            finally { Log.exit(); }
        }
        /// <summary>
        /// выполнение Процесса name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>List<int> results</int></returns>
        /// <issue> 
        /// 1.1.14 - не написана интеграция результатов отдельных Шагов в общий результат Процесса
        /// </issue>
        /// <journal> 31.12.2013
        /// </journal>
        public static List<int> Start(string name)
        {
            Log.set("Process.Start(" + name + ")");
            Process proc = Processes[name];
            //**************************************************************
            foreach (Step stp in proc.steps) stp.Exec();    //**** Step.Exec
            //**************************************************************

            // по списку Документов в строке <*>ProcEnd формируем результаты:
            // число измененных файлов и число записей в каждом из них
            List<string> docNames = new List<string>();
            //////Excel.Range rw = proc._rng.Range[proc._rng.Rows.Count];     //ПЕРЕПИСАТЬ!!
            //////docNames = MatchLib.ToStrList(rw.Range[Decl.STEP_DOCS]);    //ПЕРЕПИСАТЬ!!
//----------------- еще не дописано!!!!!!!! --------------------
//            docNames.ForEach(docAcc => docAcc = Docs.getDoc(
//            proc.results[0].Value = (int) docs.Count;
//            proc.results.ForEach(newLines => newLies = 
            Log.exit();
            return Processes[name].results;
        }
        /// <summary>
        /// Reset(name) - сбрасываем и запускаем заново Процесс name -- Start(name)
        /// </summary>
        /// <param name="name"></param>
        /// <returns>List<int> results</int></returns>
        /// <journal> 1.1.2014 PKh
        /// 3.2.14 - еще не дописан сброс окраски строк Процесса
        /// </journal>
        public static List<int> Reset(string name)
        {
            Log.set("Process.Reset(" + name + ")");
            Process proc = Processes[name];
            //-------
            // сбросить окраску строк Процесса от proc._rng[0] до proc._rng[1]
            //
            //int rngCount = proc._rng[1] - proc._rng[0];
            //////int rngCount = proc._rng.Rows.Count;
            //////proc._rng.Range["A1:C" + rngCount].Interior.Color = Excel.Constants.xlNone;
            proc.steps.ForEach(stp => stp.Reset());
            proc.results = Process.Start(name);
            Log.exit();
            return Processes[name].results;
        }
        /// <summary>
        /// Шаг - атомарная часть Процесса
        /// </summary>
        public class Step
        {
            private string name;                                // имя Шага
            private bool done;                                  // не пусто - Шаг выполнен
            private List<string> prevSteps = new List<string>();    // предыдущие Шаги
            private List<string> parameters = new List<string>();   // входные параметры Шага
            public List<string> docNames = new List<string>();      // имена Документов, обрабатываемае в Шаге
 //           private List<Docs> docs = new List<Docs>();         // собственно Документы
            private List<int> results = new List<int>();        // один или несколько Результатов Шага
            private object[] _stepRow;                       // строка таблицы Процессов по Шагу 
            /// <summary>
            /// разбор строки rw для заполнения спецификации Шага
            /// </summary>
            /// <param name="rw"></param>
            /// <journal>
            /// 23.1.14 - overridden Step(DataRow)
            /// 17.1.15 - без DataRow
            /// </journal>
            public Step(object[] rw)
            {
                name = rw[Decl.STEP_NAME] as string;
                Log.set("конструктор Step(" + name + "...)");
                done = !String.IsNullOrEmpty(rw[Decl.STEP_DONE] as string);
                _stepRow = rw;

                parameters  = MatchLib.ToStrList(rw, Decl.STEP_PARAMS);
                docNames    = MatchLib.ToStrList(rw, Decl.STEP_DOCS);
                prevSteps   = MatchLib.ToStrList(rw[Decl.STEP_PREV] as string);
                Log.exit();
            }
            /// <summary>
            /// Step.Reset()    - cброс времени и статуса Шага done
            /// </summary>
            public void Reset()
            {
                this.done = false;
                _stepRow[Decl.STEP_TIME] = "";
                _stepRow[Decl.STEP_DONE] = "";
            }
            /// <summary>
            /// Step.Exec()     - выполнение Шага
            /// </summary>
            /// <returns>results - результаты работы ШАГА</returns>
            /// <remarks>
            /// - при работе статического конструктора из листа Process извлекаются только
            ///   строки - имена Документов. При исполнении Шага с диска загружаются и
            ///   проверяются реальные Документы. Таким образом, в списке docNames вполне могут
            ///   быть пустые строки или несуществующие имена Документов; лишь при исполнении
            ///   Процесса может возникнуть сообщение об ошибке - о несуществующем Документе.
            /// </remarks>
            public void Exec()
            {
                Log.set("Log.Exec(" + name + ")");
                if (done) goto exit;               // Шаг уже выполнен -> return
                // извлечь Документы, обрабатываемые в Шаге
                //foreach (string docName in docNames)
                //    if (docName != "") docs.Add(Docs.getDoc(docName));
                // проверим необходимые Шаги контекста -- PrevStep
                if (prevSteps != null && prevSteps.Any())
                {
                    foreach (var itemPrevStep in prevSteps)
                    {
                        switch (itemPrevStep)
                        {
                            case "": break;
                            case "Loaded":
                                {
                                    //   if (this.prevSteps.Contains("Loaded"))
                                    // тут проверить ПЕРВЫЙ обрабатываемый Документ
                                    // этот Документ должен быть docAcc.MadeStep == "Loaded"
                                    // а если это не так Log.FATAL
                                }
                                break;
                            default:
                                {
                                    // структура PrevStep: Процесс/Шаг
                                    // разбор и, если нужно, исполнение Шага в звене prevStepItem
                                    // вначале отделим имя Процесса ('/'), на которое ссылается itemPrevStep
                                    // если оно опушено, то это PrevStep в текущем Процессе
                                }
                                break;
                        }
                    }
                }
                //-------------------------------------------
                // теперь передадим управление методу Шага
                // обращение имеет вид docs[0].Метод([parameters],[docs])  

                MethodInfo[] methods = typeof(match.Handler.Handler).GetMethods();

                foreach (MethodInfo info in methods)
                {
 //                   Console.WriteLine(info.Name);
                    if (info.Name == name)
                    {
                        info.Invoke(new match.Handler.Handler(parameters, docNames), null);
                        goto stepDone;
                    }
                }
                Log.FATAL("Шаг \"" + name + "\" вызывает несуществующий метод.");
                // сохраним Документ, обработанный в Шаге
stepDone:       Docs doc = Docs.getDoc(docNames[0]);
                FileOp.fileSave(doc.Wb);
                // сохраним изменения в файле match.xlsm --- НАДО ПЕРЕСМОТРЕТЬ!
                done = true;
                _stepRow[Decl.STEP_DONE] = "1";
                //_stepRow[1].Color = System.Drawing.Color.LightGreen;
                //_stepRow[2].Color = System.Drawing.Color.LightGreen;
                //_stepRow[3].Color = System.Drawing.Color.LightGreen;
                //_stepRow.Range["A1:C1"].Interior.Color = Excel.XlRgbColor.rgbLightGreen;
                Docs docProc = Docs.getDoc(Decl.PROCESS);
                docProc.isChanged = true;
//!!                Docs.saveDoc(Decl.PROCESS);
exit:           Log.exit();
            }
        }
    }
}