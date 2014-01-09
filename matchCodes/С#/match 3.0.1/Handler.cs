using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Docs = match.Document.Document;
using match.Process;
using Log = match.Lib.Log;
using Decl = match.Declaration.Declaration;

namespace match.Handler
{
    class Handler
    {
        Dictionary<string,Docs> docs = new Dictionary<string, Docs>();
 //       List<Excel.Range> patterns = new List<Excel.Range>();
      
        public Handler(List<string> parameters, List<string> docNames)
        {
            foreach (string docName in docNames)
                if (docName != "") docs.Add(docName, Docs.getDoc(docName));
        }

        /// <summary>
        /// ��������� ������� � �������� 
        /// </summary>
        /// <journal>7.1.2013 PKh</journal>
        public void InsMyCol()
        {
            Log.set("InsMyCol");
            new Log("����� �������� InsMyCol");
//            Docs doc = docs[0];
            Docs doc = docs.First().Value;

            if (doc.Body.Range["A1"].Text == doc.BodyPtrn.Range["A1"].Text)
                Log.FATAL("������� ���������� ��� ������������ ��������");
            //---  ��������� ������� �� ����� MyCol        
            doc.Sheet.Range["A1", doc.Body.Cells[1, doc.MyCol]].EntireColumn.Insert();
            //--- ������������� ������ ������� ����� �� ��������� � ������ ������� Width
            new Log("����� ������� �������");
            int i = 1;
            foreach (Excel.Range col in doc.BodyPtrn.Columns)
            {
                string s = col.Range[Decl.PTRN_WIDTH].Text;
                if (s == Decl.PTRN_COPYHDR) col.Range["A1"].Copy(doc.Body.Cells[1, i]);
                string[] ar = s.Split('/');
                float W;
                if (!float.TryParse(ar[0], out W)) Log.FATAL("������ � ������ Width ������� = \""
                    + s + "\" ��� ��������� ��������� " + doc.name); 
                doc.Body.Columns[i++].ColumnWidth = W;
            }
            //--- �������� ������� MyCol �� ����� �� Body.EOL
            doc.BodyPtrn.Range["A1", doc.BodyPtrn.Cells[2, doc.MyCol]].Copy(doc.Body.Range["A1"]);
            doc.Body.Range["A2", doc.Body.Cells[doc.Body.Rows.Count, doc.MyCol]].FillDown();
            //--- ���� ���� --> ��������� �����
            if (doc.SummaryPtrn != null) doc.SummaryPtrn.Copy(doc.Summary.Range["A2"]);
            Log.exit();
        }
        public void DateSort()
        {
        }
        public void PaymentPaint()
        {
        }
        public void ContractPaint()
        {
        }
        public void SF_Paint()
        {
        }
        public void AccPaint()
        {
        }
        public void Acc1C_Bottom()
        {
        }
        public void DicAccSyn()
        {
            const string SF_ACC_SYNONIMS = "SF_DicAccSyn";
            const string DOC_ACC_SYNONIMS = "DicAccSynonims";
            const string SYN_VALUE_COL = "C1";  // ������� 2 - ������ ���������
            string[] ACC_DEL = { "<���>" };

            Log.set("DicAccSyn");
            try
            {
                Docs docSF  = docs[SF_ACC_SYNONIMS];
                Docs doc    = docs[DOC_ACC_SYNONIMS];
                doc.Reset();
                Excel.Range Bdy = doc.Body;

                //      ���� �� ���� ������� �����
                int rowNum = 2;
                foreach (Excel.Range row in docSF.Body.Rows)
                {
                    string[] syn = ((string)row.Range[SYN_VALUE_COL].Text)
                        .Split(ACC_DEL, StringSplitOptions.RemoveEmptyEntries);
                    if (syn.Length < 2) continue;
                    foreach (string str in syn)
                    {
//                        Excel.Range rw = doc.AddRow();
                        doc.Body.Range["A" + rowNum].Value = str.Trim();
                        doc.Body.Range["B" + rowNum].Value = row.Range[SYN_VALUE_COL].Text;
                        rowNum++;
                    }
                }
            }
            finally { Log.exit(); }
        }
        public void RowDel()
        {
        }
        public void CheckRepDate()
        {
        }
        public void MergeReps()
        {
        }
        public void Adapt()
        {

            Log.set("Adapt");
            try
            {
                Excel.Workbook db_match = match.MyFile.FileOpenEvent.fileOpen(Decl.F_MATCH); //
                Excel._Worksheet hdrSht = db_match.Worksheets[Decl.HEADER];
                Excel.Range ptrn;
                try { ptrn = hdrSht.get_Range("HDR_??"); }  catch { ptrn = null; }

            }
            finally
            {
                Log.exit();
            }
        }
        public void ProcStart()
        {
        }
        public void Paid1C()
        {
        }
        public void WrCSV()
        {
        }
    }
}
