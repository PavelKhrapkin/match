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
        /// вставл€ет колонки в ƒокумент 
        /// </summary>
        /// <journal>7.1.2013 PKh</journal>
        public void InsMyCol()
        {
            Log.set("InsMyCol");
            new Log("перед запуском InsMyCol");
//            Docs doc = docs[0];
            Docs doc = docs.First().Value;

            if (doc.Body.Range["A1"].Text == doc.BodyPtrn.Range["A1"].Text)
                Log.FATAL("ѕопытка обработать уже обработанный ƒокумент");
            //---  вставл€ем колонки по числу MyCol        
            doc.Sheet.Range["A1", doc.Body.Cells[1, doc.MyCol]].EntireColumn.Insert();
            //--- устанавливает ширину колонки листа по значени€м в строке Ўаблона Width
            new Log("после вставки колонок");
            int i = 1;
            foreach (Excel.Range col in doc.BodyPtrn.Columns)
            {
                string s = col.Range[Decl.PTRN_WIDTH].Text;
                if (s == Decl.PTRN_COPYHDR) col.Range["A1"].Copy(doc.Body.Cells[1, i]);
                string[] ar = s.Split('/');
                float W;
                if (!float.TryParse(ar[0], out W)) Log.FATAL("ошибка в строке Width шаблона = \""
                    + s + "\" при обработке ƒокумента " + doc.name); 
                doc.Body.Columns[i++].ColumnWidth = W;
            }
            //--- копируем колонки MyCol от верха до Body.EOL
            doc.BodyPtrn.Range["A1", doc.BodyPtrn.Cells[2, doc.MyCol]].Copy(doc.Body.Range["A1"]);
            doc.Body.Range["A2", doc.Body.Cells[doc.Body.Rows.Count, doc.MyCol]].FillDown();
            //--- если есть --> формируем п€тку
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
            const string SYN_VALUE_COL = "C1";  // колонка 2 - список синонимов
            string[] ACC_DEL = { "<»Ћ»>" };

            Log.set("DicAccSyn");
            try
            {
                Docs docSF  = docs[SF_ACC_SYNONIMS];
                Docs doc    = docs[DOC_ACC_SYNONIMS];
                doc.Reset();
                Excel.Range Bdy = doc.Body;

                //      цикл по всем строкам листа
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
