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
    /// ����� Process (�������) �������� ������� ���� ��������� ������� ���������
    /// </summary>
    /// <journal> 29.12.2013
    /// 30.12.2013 - ���������� � C# �� match 2.2 ������ ProcessEngin.bas
    /// 31.12.2013 -- �������� ���������� ���������� ProccessEngine
    /// </journal>
    class Process
    {
        /// <summary>
        /// Processes - ��������� ���������
        /// </summary>
        private static Dictionary<string, Process> Processes = new Dictionary<string, Process>();

        private string name;                            // ��� ��������
        private List<Step> steps = new List<Step>();    // ����, ������������ �������
        public List<int> results = new List<int>();     // ���� ��� ��������� ����������� ��������
        private Excel.Range _rng;
        private int line_procStart;

        /// <summary>
        /// ����������� ����������� �������� ����������� �������� ����� ������ ���������� � ������.
        /// �� ��������� ������ � ��������� Processes �� ����� Process ����� match.xlsm
        /// </summary>
        /// <journal> 1.1.2014
        /// </journal>
        static Process()
        {
            Log.set("����������� ����������� ���������");
            try
            {
                Process proc = null;
                Docs proc_doc = Docs.getDoc("Process");
                int line_process = 0;
                foreach (Excel.Range rw in proc_doc.Body.Rows)
                {
                    line_process++;
                    var cell = rw.Range[Decl.STEP_NAME].Value2;
                    if (cell == null || rw.Range[Decl.STEP_COMMENT].Value2 != null) continue;
                    switch ((string)cell.Text)
                    {
                        case "<*>ProcStart":
                            proc = new Process();
                            proc.line_procStart = line_process;
                            proc.name = rw.Range[Decl.PROC_NAME].Text;
                            break;
                        case "<*>ProcEnd":
                            proc._rng = proc_doc.Body.Range[proc.line_procStart + ":" + line_process];
                            Processes.Add(proc.name, proc);
                            break;
                        default:
                            if (proc.name == null) continue;    //������ ���
                            proc.steps.Add(new Step(rw));
                            break;
                    }
                }
            }
            catch
            {
                Log.FATAL("�� ������� ���������������� ������� ���������");
            }
            Log.exit();
        }
        /// <summary>
        /// ���������� �������� name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>List<int> results</int></returns>
        /// <issue> 
        /// 1.1.14 - �� �������� ���������� ����������� ��������� ����� � ����� ��������� ��������
        /// </issue>
        /// <journal> 31.12.2013
        /// </journal>
        public static List<int> Run(string name)
        {
            Log.set("Process.Run(" + name + ")");
            Process proc = Processes[name];
            //**************************************************************
            foreach (Step stp in proc.steps) stp.Exec();    //**** Step.Exec
            //**************************************************************
            Log.exit();
            return Processes[name].results;
        }
        /// <summary>
        /// Reset(name) - ���������� � ��������� ������ ������� name -- Run(name)
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
        /// ��� - ��������� ����� ��������
        /// </summary>
        private class Step
        {
            private string name;                                // ��� ����
            private bool done;                                  // �� ����� - ��� ��������
            private List<string> prevSteps = new List<string>();    // ���������� ����
            private List<string> parameters = new List<string>();   // ������� ��������� ����
            private List<string> docNames = new List<string>();     // ����� ����������, �������������� � ����
            private List<Docs> docs = new List<Docs>();         // ���������� ���������
            private List<int> results = new List<int>();        // ���� ��� ��������� ����������� ����
            private Excel.Range _stepRow;                       // ������ ������� ��������� �� ���� 
            /// <summary>
            /// ������ ������ rw ��� ���������� ������������ ����
            /// </summary>
            /// <param name="rw"></param>
            public Step(Excel.Range rw)
            {
                name = rw.Range[Decl.STEP_NAME].Text;
                Log.set("����������� Step(" + name + "...)"); 
                done = !String.IsNullOrEmpty(rw.Range[Decl.STEP_DONE].Text);
                _stepRow = rw;

                parameters  = MatchLib.ToStrList(rw.Range[Decl.STEP_PARAMS]);
                docNames    = MatchLib.ToStrList(rw.Range[Decl.STEP_DOCS]);
                prevSteps   = MatchLib.ToStrList(rw.Range[Decl.STEP_PREV].Text);
            }
            /// <summary>
            /// Step.Reset()    - c���� ������� � ������� ���� done
            /// </summary>
            public void Reset()
            {
                this.done = false;
                _stepRow.Range[Decl.STEP_TIME].Value = "";
                _stepRow.Range[Decl.STEP_DONE].Value = "";
            }
            /// <summary>
            /// Step.Exec()     - ���������� ����
            /// </summary>
            /// <returns>results - ���������� ������ ����</returns>
            /// <remarks>
            /// - ��� ������ ������������ ������������ �� ����� Process ����������� ������
            ///   ������ - ����� ����������. ��� ���������� ���� � ����� ����������� �
            ///   ����������� �������� ���������. ����� �������, � ������ docNames ������ �����
            ///   ���� ������ ������ ��� �������������� ����� ����������; ���� ��� ����������
            ///   �������� ����� ���������� ��������� �� ������ - � �������������� ���������.
            /// </remarks>
            public List<int> Exec()
            {
                Log.set("Log.Exec(" + name + ")");
                if (!done)                // ��� ��� �������� -> return
                {
                    // ������� ���������, �������������� � ����
                    foreach (string docName in docNames)
                        if (docName != "") docs.Add(Docs.getDoc(docName));
                    // �������� ����������� ���� ��������� -- PrevStep
                    foreach (var itemPrevStep in prevSteps)
                    {
                        switch (itemPrevStep)
                        {
                            case "": break;
                            case "Loaded":
                                {
                                    //   if (this.prevSteps.Contains("Loaded"))
                                    // ��� ��������� ������ �������������� ��������
                                    // ���� �������� ������ ���� doc.MadeStep == "Loaded"
                                    // � ���� �� �� ��� Log.FATAL
                                }
                                break;
                            default:
                                {
                                    // ��������� PrevStep: �������/���
                                    // ������ �, ���� �����, ���������� ���� � ����� prevStepItem
                                    // ������� ������� ��� �������� ('/'), �� ������� ��������� itemPrevStep
                                    // ���� ��� �������, �� ��� PrevStep � ������� ��������
                                }
                                break;
                        }
                    }
                    //-------------------------------------------
                    // ������ ��������� ���������� ������ ����
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