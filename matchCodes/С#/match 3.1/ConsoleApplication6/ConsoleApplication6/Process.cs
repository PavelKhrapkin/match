/*-----------------------------------------------------------------------
 * Process -- ����� ��������� ������� match 3.1
 * 
 *  17.01.2015  �.�������, �.����
 *  
 * -------------------------------------------
 * Process()            - ����������� ��������� ������� ���������
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
    /// ����� Process (�������) �������� ������� ���� ��������� ������� ���������
    /// </summary>
    /// <journal> 29.12.2013
    /// 30.12.2013 - ���������� � C# �� match 2.2 ������ ProcessEngin.bas
    /// 31.12.2013 -- �������� ���������� ���������� ProccessEngine
    /// 17.1.2015 - ���������� ��� Data Table
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
        private int[] _rng;                             //������ ����� <*>ProcStart � <*>ProcEnd ��������
        private int line_procStart;

        /// <summary>
        /// ����������� ����������� �������� ����������� �������� ����� ������ ���������� � ������.
        /// �� ��������� ������ � ��������� Processes �� ����� Process ����� match.xlsm
        /// </summary>
        /// <journal> 1.1.2014
        /// 17.1.2015 - ���������� ��� Data Table
        /// </journal>
        static Process()
        {
            Log.set("����������� ����������� ���������");
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
                            if (proc == null || proc.name == null) continue;    //������ ���
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
                            if (proc.name == null) continue;    //������ ���
                            proc.steps.Add(new Step(rw));
                            break;
                    }
                }
 PK */
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
            //////Excel.Range rw = proc._rng.Range[proc._rng.Rows.Count];     //����������!!
            //////docNames = MatchLib.ToStrList(rw.Range[Decl.STEP_DOCS]);    //����������!!
//----------------- ��� �� ��������!!!!!!!! --------------------
//            docNames.ForEach(docAcc => docAcc = Docs.getDoc(
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
        /// 3.2.14 - ��� �� ������� ����� ������� ����� ��������
        /// </journal>
        public static List<int> Reset(string name)
        {
            Log.set("Process.Reset(" + name + ")");
            Process proc = Processes[name];
            //-------
            // �������� ������� ����� �������� �� proc._rng[0] �� proc._rng[1]
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
        /// ��� - ��������� ����� ��������
        /// </summary>
        public class Step
        {
            private string name;                                // ��� ����
            private bool done;                                  // �� ����� - ��� ��������
            private List<string> prevSteps = new List<string>();    // ���������� ����
            private List<string> parameters = new List<string>();   // ������� ��������� ����
            public List<string> docNames = new List<string>();      // ����� ����������, �������������� � ����
 //           private List<Docs> docs = new List<Docs>();         // ���������� ���������
            private List<int> results = new List<int>();        // ���� ��� ��������� ����������� ����
            private object[] _stepRow;                       // ������ ������� ��������� �� ���� 
            /// <summary>
            /// ������ ������ rw ��� ���������� ������������ ����
            /// </summary>
            /// <param name="rw"></param>
            /// <journal>
            /// 23.1.14 - overridden Step(DataRow)
            /// 17.1.15 - ��� DataRow
            /// </journal>
            public Step(object[] rw)
            {
                name = rw[Decl.STEP_NAME] as string;
                Log.set("����������� Step(" + name + "...)");
                done = !String.IsNullOrEmpty(rw[Decl.STEP_DONE] as string);
                _stepRow = rw;

                parameters  = MatchLib.ToStrList(rw, Decl.STEP_PARAMS);
                docNames    = MatchLib.ToStrList(rw, Decl.STEP_DOCS);
                prevSteps   = MatchLib.ToStrList(rw[Decl.STEP_PREV] as string);
                Log.exit();
            }
            /// <summary>
            /// Step.Reset()    - c���� ������� � ������� ���� done
            /// </summary>
            public void Reset()
            {
                this.done = false;
                _stepRow[Decl.STEP_TIME] = "";
                _stepRow[Decl.STEP_DONE] = "";
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
                                    // ��� ��������� ������ �������������� ��������
                                    // ���� �������� ������ ���� docAcc.MadeStep == "Loaded"
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
                }
                //-------------------------------------------
                // ������ ��������� ���������� ������ ����
                // ��������� ����� ��� docs[0].�����([parameters],[docs])  

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
                Log.FATAL("��� \"" + name + "\" �������� �������������� �����.");
                // �������� ��������, ������������ � ����
stepDone:       Docs doc = Docs.getDoc(docNames[0]);
                FileOp.fileSave(doc.Wb);
                // �������� ��������� � ����� match.xlsm --- ���� ������������!
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