using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using match.MyFile;
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
        private Excel.Range _rng;
        private int line_procStart;

        /// <summary>
        /// Статический конструктор Процесса запускается системой перед первым обращением к классу.
        /// Он считывает данные в коллекцию Processes из листа Process файла match.xlsm
        /// </summary>
        /// <journal> 1.1.2014
        /// </journal>
        static Process()
        {
            Log.set("статический конструктор Процессов");
            try {
                Process proc = null;
                Docs proc_doc = Docs.getDoc(Decl.PROCESS);
                int line_process = 0;
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
            Excel.Range rw = proc._rng.Range[proc._rng.Rows.Count];
            docNames = MatchLib.ToStrList(rw.Range[Decl.STEP_DOCS]);
//----------------- еще не дописано!!!!!!!! --------------------
//            docNames.ForEach(doc => doc = Docs.getDoc(
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
        /// </journal>
        public static List<int> Reset(string name)
        {
            Log.set("Process.Reset(" + name + ")");
            Process proc = Processes[name];
            int rngCount = proc._rng.Rows.Count;
            proc._rng.Range["A1:C" + rngCount].Interior.Color = Excel.Constants.xlNone;
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
            public List<string> docNames = new List<string>();     // имена Документов, обрабатываемае в Шаге
 //           private List<Docs> docs = new List<Docs>();         // собственно Документы
            private List<int> results = new List<int>();        // один или несколько Результатов Шага
            private Excel.Range _stepRow;                       // строка таблицы Процессов по Шагу 
            /// <summary>
            /// разбор строки rw для заполнения спецификации Шага
            /// </summary>
            /// <param name="rw"></param>
            public Step(Excel.Range rw)
            {
                name = rw.Range[Decl.STEP_NAME].Text;
                Log.set("конструктор Step(" + name + "...)"); 
                done = !String.IsNullOrEmpty(rw.Range[Decl.STEP_DONE].Text);
                _stepRow = rw;

                parameters  = MatchLib.ToStrList(rw.Range[Decl.STEP_PARAMS]);
                docNames    = MatchLib.ToStrList(rw.Range[Decl.STEP_DOCS]);
                prevSteps   = MatchLib.ToStrList(rw.Range[Decl.STEP_PREV].Text);
                Log.exit();
            }
            /// <summary>
            /// Step.Reset()    - cброс времени и статуса Шага done
            /// </summary>
            public void Reset()
            {
                this.done = false;
                _stepRow.Range[Decl.STEP_TIME].Value = "";
                _stepRow.Range[Decl.STEP_DONE].Value = "";
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
                foreach (var itemPrevStep in prevSteps)
                {
                    switch (itemPrevStep)
                    {
                        case "": break;
                        case "Loaded":
                            {
                                //   if (this.prevSteps.Contains("Loaded"))
                                // тут проверить ПЕРВЫЙ обрабатываемый Документ
                                // этот Документ должен быть doc.MadeStep == "Loaded"
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
                //-------------------------------------------
                // теперь передадим управление методу Шага
                // обращение имеет вид docs[0].Метод([parameters],[docs])  

                MethodInfo[] methods = typeof(match.Handler.Handler).GetMethods();

                foreach (MethodInfo info in methods)
                {
                    Console.WriteLine(info.Name);
                    if (info.Name == name)
                    {
                        info.Invoke(new match.Handler.Handler(parameters, docNames), null);
                        goto stepDone;
                    }
                }
                Log.FATAL("Шаг \"" + name + "\" вызывает несуществующий метод.");
                // сохраним Документ, обработанный в Шаге
stepDone:       Docs doc = Docs.getDoc(docNames[0]);
                FileOpenEvent.fileSave(doc.Wb);
                // сохраним изменения в таблице Process
                done = true;
                _stepRow.Range[Decl.STEP_DONE].Value2 = "1";
                _stepRow.Range["A1:C1"].Interior.Color = Excel.XlRgbColor.rgbLightGreen;
                Docs docProc = Docs.getDoc(Decl.PROCESS);
                docProc.isChanged = true;
                Docs.saveDoc(Decl.PROCESS);
exit:           Log.exit();
            }
        }
    }
}