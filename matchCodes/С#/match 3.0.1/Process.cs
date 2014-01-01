using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using Decl = match.Declaration.Declaration;
using Docs = match.Document.Document;
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
        private List<string> docNames = new List<string>();
        private List<Docs> docs = new List<Docs>();     // Документы, обрабатываемае в Процессе
        public List<int> results = new List<int>();     // один или несколько Результатов Процесса
        private Excel.Range _rng;
        private int line_procStart;

        /// <summary>
        /// Статический конструктор Процесса запускается системой перед первым обращением к классу.
        /// Он считывает данные в коллекцию Processes из листа Process файла match.xlsm
        /// </summary>
        static Process()
        {
            Log.set("статический конструктор Процессов");
            try
            {
                Process proc = null;
                Docs proc_doc = Docs.getDoc("Process");
                string StepName;
                int line_process = 0;
                foreach (Excel.Range rw in proc_doc.Body.Rows)
                {
                    line_process++;
                    var cell = rw.Range[Decl.STEP_NAME].Value2;
                    if (cell == null || rw.Range[Decl.STEP_COMMENT].Value2 != null) continue;
                    StepName = cell.ToString();

                    switch (StepName)
                    {
                        case "<*>ProcStart":
                            proc = new Process();
                            proc.line_procStart = line_process;
                            proc.name = rw.Range[Decl.PROC_NAME].Text;
                            //загрузим имена обрабатываемых в Процессе Документов
                            Excel.Range rng_docs = rw.Range[Decl.PROC_DOCUMENTS];
                            proc.docNames = MatchLib.ToStrList(rng_docs);
                            break;
                        case "<*>ProcEnd":
                            proc._rng = proc_doc.Body.Range[proc.line_procStart + ":" + line_process];
                            Processes.Add(proc.name, proc);
                            StepName = null;
                            break;
                        default:
                            if (proc.name == null) continue;    //пустой Шаг
                            // добавим Шаг в Процесс StepName
                            Step stp = new Step(rw);
                            proc.steps.Add(stp);
                            break;
                    }
                }
            }
            catch
            {
                Log.FATAL("Не удалось инициализировать таблицу Процессов");
            }
            Log.exit();
        }
        /// <summary>
        /// выполнение Процесса name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>List<int> results</int></returns>
        /// <remarks>
        /// - при работе статического конструктора из листа Process извлекаются только
        ///   строки - имена Документов. При исполнении Процесса с диска загружаются и
        ///   проверяются реальные Документы. Таким образом, в списке docNames вполне могут
        ///   быть пустые строки или несуществующие имена Документов; лишь при исполнении
        ///   Процесса может возникнуть сообщение об ошибке - о несуществующем Документе.
        /// </remarks>
        /// <journal> 31.12.2013
        /// </journal>
        public static List<int> Run(string name)
        {
            Log.set("Process.Run(" + name + ")");
            Process proc = Processes[name];
            // извлечь Документы, обрабатываемые Процессом
            foreach (string docName in proc.docNames)
                if (docName != "") proc.docs.Add(Docs.getDoc(docName));
            foreach (Step stp in proc.steps)
                stp.Exec();
            Log.exit();
            return Processes[name].results;
        }
        /// <summary>
        /// Reset(name) - сбрасываем и запускаем заново Процесс name -- Run(name)
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
            proc.results = Process.Run(name);
            Log.exit();
            return Processes[name].results;
        }
        /// <summary>
        /// Шаг - атомарная часть Процесса
        /// </summary>
        private class Step
        {
            private string name;                                // имя Шага
            private bool done;                                  // не пусто - Шаг выполнен
            private List<string> prevSteps = new List<string>();    // предыдущие Шаги
            private List<string> parameters = new List<string>();   // входные параметры Шага
            private List<string> docs = new List<string>();     // имена Документов, обрабатываемае в Шаге
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

                parameters =    MatchLib.ToStrList(rw.Range[Decl.STEP_PARAMS]);
                docs =          MatchLib.ToStrList(rw.Range[Decl.STEP_DOCS]);
                prevSteps =  MatchLib.ToStrList(rw.Range[Decl.STEP_PREV].Text);

                /* ----------- все это можно будет стереть -----------
                //// разбор параметров и Документов Шага
                //string[]  step_pars = {"F1", "G1", "H1", "I1", "J1"};
                //foreach (var cell in step_pars) parameters.Add(rw.Range[cell].Text);

                //string[] step_docs = { "K1", "L1", "M1", "N1", "O1" };
                //foreach (var cell in step_docs) docs.Add(rw.Range[cell].Text);
                
                //List<string> fiveCell = new List<string>() {"F1", "G1", "H1", "I1", "J1"};
                //fiveCell.ForEach(cell=> parameters.Add(rw.Range[cell].Text));
                 ------------------------------------------------------- */
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
            /// выполнение Шага
            /// </summary>
            /// <returns></returns>
            public List<int> Exec()
            {
                Log.set("Log.Exec(" + this.name + ")");
                if (!this.done)                // Шаг уже выполнен -> return
                {
                    // проверим необходимые Шаги контекста -- PrevStep
                    foreach (var itemPrevStep in this.prevSteps)
                    {
                        switch (itemPrevStep)
                        {
                            case "": break;
                            case "Loaded":
                                {
                                    //   if (this.prevSteps.Contains("Loaded"))
                                    // тут проверить основной обрабатываемый Документ
                                    //(КАКОЙ?? он должен быть определен в строке <*>ProcStart)
                                    //этот Документ должен быть doc.MadeStep == "Loaded"
                                    // а если то не так Log.FATAL
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

                    //this.prevSteps.ForEach(stp => (stp.name == "Loaded")? continue: stp.Exec());
                    //this.docs.ForEach(doc => 
                    //    {
                    //        Docs.getDoc();
                    //    }
                    Program program = new Program();
                    MethodInfo[] methods = typeof(Process).GetMethods();
                    foreach (MethodInfo info in methods)
                    {
                        Console.WriteLine(info.Name);

                        // Call Win method.
                        if (info.Name == "Win")
                        {
                            info.Invoke(program, null); // [2]
                        }
                    }
                }
                Log.exit();
                return results;
            }
        }
    }
}