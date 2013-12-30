using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using Decl = match.Declaration.Declaration;
using Docs = match.Document.Document;
using Lib = match.Lib;
using Log = match.Lib.Log;

namespace match.Process
{
    /// <summary>
    /// класс Process (Процесс) содержит таблицу всех известных системе Процессов
    /// </summary>
    /// <journal> 29.12.2013
    /// 30.12.2013 - перенесено в C# из match 2.2 модуля ProcessEngin.bas
    /// </journal>
    class Process
    {
        /// <summary>
        /// Processes - коллекция Процессов
        /// </summary>
        private static Dictionary<string, Process> Processes = new Dictionary<string, Process>();

        private string name;                            // имя Процесса
        private List<Step> steps = new List<Step>();    // Шаги, составляющие Процесс
        private List<Docs> docs = new List<Docs>();     // Документы, обрабатываемае в Процессе
        public List<int> results = new List<int>();     // один или несколько Результатов Процесса

        /// <summary>
        /// Статический конструктор Процесса запускается системой перед первым обращением к классу.
        /// Он считывает данные в коллекцию Processes из листа Process файла match.xlsm
        /// </summary>
        static Process()
        {
            Log.set("статический конструктор Процессов");
            Process proc = null;
            Docs proc_doc = Docs.getDoc("Process");
            string StepName;
            foreach (Excel.Range rw in proc_doc.Body.Rows)
            {
                var cell = rw.Range[Decl.STEP_NAME].Value2;
                if (cell == null || rw.Range[Decl.STEP_COMMENT].Value2 != null) continue;
                StepName = cell.ToString();

                switch (StepName)
                {
                    case "<*>ProcStart":
                        proc = new Process();
                        proc.name = rw.Range[Decl.PROC_NAME].Value2.ToString();
                        break;
                    case "<*>ProcEnd":
                        Processes.Add(proc.name, proc);
                        StepName = null;
                        break;
                    default:
                        if (proc.name == null) continue;    //пустой Шаг
                    // тут надо добавить Шаг в Процесс StepName
                        Step stp = new Step(rw);
                        proc.steps.Add(stp);
                        break;
                }
            }
            Log.exit();
        }

        public static int Run(string name)
        {
            Log.set("Process.Run(" + name + ")");

            Log.exit();
            return 0;
        }
        public void Exec(string name)
        {
        }
        /// <summary>
        /// Шаг - атомарная часть Процесса
        /// </summary>
        private class Step
        {
            private string name;                                // имя Шага
            private bool done;                                  // не пусто - Шаг выполнен
            private List<Step> prevSteps = new List<Step>();    // обязательно выполненные предыдущие Шаги
            private List<string> parameters = new List<string>();   // входные параметры Шага
            private List<Docs> docs = new List<Docs>();         // Документы, обрабатываемае в Шаге
            public List<int> results = new List<int>();         // один или несколько Результатов Шага

            public Step(Excel.Range rw)
            {
                // разбор строки rw для заполнения спецификации Шага
                name = rw.Range[Decl.STEP_NAME].Text;
                done = !String.IsNullOrEmpty(rw.Range[Decl.STEP_DONE].Text);

                // разбор параметров и Документов Шага
                string[]  step_pars = {"F1", "G1", "H1", "I1", "J1"};
                foreach (var cell in step_pars) parameters.Add(rw.Range[cell].Text);

                string[] step_docs = { "K1", "L1", "M1", "N1", "O1" };
                foreach (var cell in step_docs) docs.Add(rw.Range[cell].Text);
                
                //List<string> fiveCell = new List<string>() {"F1", "G1", "H1", "I1", "J1"};
                //fiveCell.ForEach(cell=> parameters.Add(rw.Range[cell].Text));
            }
        }
    }
}