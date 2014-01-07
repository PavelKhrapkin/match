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
        List<Docs> docs = new List<Docs>();
 //       List<Excel.Range> patterns = new List<Excel.Range>();
      
        public Handler(List<string> parameters, List<string> docNames)
        {
            foreach (string docName in docNames)
                if (docName != "") docs.Add(Docs.getDoc(docName));
        }

        /// <summary>
        /// ��������� ������� � �������� 
        /// </summary>
        /// <journal>7.1.2013 PKh</journal>
        public void InsMyCol()
        {
            Log.set("InsMyCol");
            new Log("����� �������� InsMyCol");
            Docs doc = docs[0];

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
        public void RowDel()
        {
        }
    }
}
