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
        public List<int> results = new List<int>();     // ���������� ��������
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
                            if (proc.name == null) continue;    //������ ���
                            proc.steps.Add(new Step(rw));
                            break;
                    }
                }
                new Log("����� ������ ������������ ������������ ���������");
            }
            catch { Log.FATAL("�� ������� ���������������� ������� ���������."); }
            finally { Log.exit(); }
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
        public static List<int> Start(string name)
        {
            Log.set("Process.Start(" + name + ")");
            Process proc = Processes[name];
            //**************************************************************
            foreach (Step stp in proc.steps) stp.Exec();    //**** Step.Exec
            //**************************************************************

            // �� ������ ���������� � ������ <*>ProcEnd ��������� ����������:
            // ����� ���������� ������ � ����� ������� � ������ �� ���
            List<string> docNames = new List<string>();
            Excel.Range rw = proc._rng.Range[proc._rng.Rows.Count];
            docNames = MatchLib.ToStrList(rw.Range[Decl.STEP_DOCS]);
//----------------- ��� �� ��������!!!!!!!! --------------------
//            docNames.ForEach(doc => doc = Docs.getDoc(
//            proc.results[0].Value = (int) docs.Count;
//            proc.results.ForEach(newLines => newLies = 
            Log.exit();
            return Processes[name].results;
        }
        /// <summary>
        /// Reset(name) - ���������� � ��������� ������ ������� name -- Start(name)
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
        /// ��� - ��������� ����� ��������
        /// </summary>
        public class Step
        {
            private string name;                                // ��� ����
            private bool done;                                  // �� ����� - ��� ��������
            private List<string> prevSteps = new List<string>();    // ���������� ����
            private List<string> parameters = new List<string>();   // ������� ��������� ����
            public List<string> docNames = new List<string>();     // ����� ����������, �������������� � ����
 //           private List<Docs> docs = new List<Docs>();         // ���������� ���������
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
                Log.exit();
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
            public void Exec()
            {
                Log.set("Log.Exec(" + name + ")");
                if (done) goto exit;               // ��� ��� �������� -> return
                // ������� ���������, �������������� � ����
                //foreach (string docName in docNames)
                //    if (docName != "") docs.Add(Docs.getDoc(docName));
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
                                // � ���� ��� �� ��� Log.FATAL
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
                // ��������� ����� ��� docs[0].�����([parameters],[docs])  

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
                Log.FATAL("��� \"" + name + "\" �������� �������������� �����.");
                // �������� ��������, ������������ � ����
stepDone:       Docs doc = Docs.getDoc(docNames[0]);
                FileOpenEvent.fileSave(doc.Wb);
                // �������� ��������� � ������� Process
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