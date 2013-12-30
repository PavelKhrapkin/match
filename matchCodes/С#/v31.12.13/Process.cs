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
    /// ����� Process (�������) �������� ������� ���� ��������� ������� ���������
    /// </summary>
    /// <journal> 29.12.2013
    /// 30.12.2013 - ���������� � C# �� match 2.2 ������ ProcessEngin.bas
    /// </journal>
    class Process
    {
        /// <summary>
        /// Processes - ��������� ���������
        /// </summary>
        private static Dictionary<string, Process> Processes = new Dictionary<string, Process>();

        private string name;                            // ��� ��������
        private List<Step> steps = new List<Step>();    // ����, ������������ �������
        private List<Docs> docs = new List<Docs>();     // ���������, �������������� � ��������
        public List<int> results = new List<int>();     // ���� ��� ��������� ����������� ��������

        /// <summary>
        /// ����������� ����������� �������� ����������� �������� ����� ������ ���������� � ������.
        /// �� ��������� ������ � ��������� Processes �� ����� Process ����� match.xlsm
        /// </summary>
        static Process()
        {
            Log.set("����������� ����������� ���������");
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
                        if (proc.name == null) continue;    //������ ���
                    // ��� ���� �������� ��� � ������� StepName
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
        /// ��� - ��������� ����� ��������
        /// </summary>
        private class Step
        {
            private string name;                                // ��� ����
            private bool done;                                  // �� ����� - ��� ��������
            private List<Step> prevSteps = new List<Step>();    // ����������� ����������� ���������� ����
            private List<string> parameters = new List<string>();   // ������� ��������� ����
            private List<Docs> docs = new List<Docs>();         // ���������, �������������� � ����
            public List<int> results = new List<int>();         // ���� ��� ��������� ����������� ����

            public Step(Excel.Range rw)
            {
                // ������ ������ rw ��� ���������� ������������ ����
                name = rw.Range[Decl.STEP_NAME].Text;
                done = !String.IsNullOrEmpty(rw.Range[Decl.STEP_DONE].Text);

                // ������ ���������� � ���������� ����
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