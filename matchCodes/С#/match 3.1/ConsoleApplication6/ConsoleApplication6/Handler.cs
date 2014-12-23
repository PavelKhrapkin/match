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
        /// вставляет колонки в Документ 
        /// </summary>
        /// <journal>7.1.2013 PKh</journal>
        public void InsMyCol()
        {
            Log.set("InsMyCol");
            new Log("перед запуском InsMyCol");
            Docs doc = docs.First().Value;
            if ((string)doc.dt.Rows[1][1] == doc.ptrn.String(1,1))
                Log.FATAL("Попытка обработать уже обработанный Документ");
/* PK
            if (doc.Body.Range["A1"].Text == doc.ptrn.Range["A1"].Text)
                Log.FATAL("Попытка обработать уже обработанный Документ");
            //---  вставляем колонки по числу MyCol        
            doc.Sheet.Range["A1", doc.Body.Cells[1, doc.MyCol]].EntireColumn.Insert();
            //--- устанавливает ширину колонки листа по значениям в строке Шаблона Width
            new Log("после вставки колонок");
            int i = 1;
            foreach (Excel.Range col in doc.ptrn.Columns)
            {
                string s = col.Range[Decl.PTRN_WIDTH].Text;
                if (s == Decl.PTRN_COPYHDR) col.Range["A1"].Copy(doc.Body.Cells[1, i]);
                string[] ar = s.Split('/');
                float W;
                if (!float.TryParse(ar[0], out W)) Log.FATAL("ошибка в строке Width шаблона = \""
                    + s + "\" при обработке Документа " + doc.name); 
                doc.Body.Columns[i++].ColumnWidth = W;
            }
            //--- копируем колонки MyCol от верха до Body.EOL
            doc.ptrn.Range["A1", doc.ptrn.Cells[2, doc.MyCol]].Copy(doc.Body.Range["A1"]);
            doc.Body.Range["A2", doc.Body.Cells[doc.Body.Rows.Count, doc.MyCol]].FillDown();
            //--- если есть --> формируем пятку
            if (doc.SummaryPtrn != null) doc.SummaryPtrn.Copy(doc.Summary.Range["A2"]);
 PK */
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

            string[] ACC_DEL = { "<ИЛИ>" };
/* PK
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
                    // извлекаем и разделяем синонимы делимитером ACC_DEL ("<ИЛИ>")
                    string[] syn = ((string)row.Range[SYN_VALUE_COL].Text)
                        .Split(ACC_DEL, StringSplitOptions.RemoveEmptyEntries);
                    if (syn.Length < 2) continue;
                    // цикл по синонимам - порождаем по строке на синоним
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
PK */ 
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

        enum pass { first, second } ;       // описание типа pass (перечисление проходов)
        public void Adapt()
        {
/* PK
            const string PTRN_TITLE = "A1";
            const string PTRN_VALUE = "A2";
            const string PTRN_WIDTH = "A3";
            const string PTRN_COLS  = "A4";
            const string PTRN_ADAPTER = "A5";
            const string PTRN_FETCH = "A6";
//            const string PTRN_1STPASS = "A7";
            Log.set("Adapt");
            try
            {
                Docs doc = docs.First().Value;
                int iRow = 0;
                foreach (Excel.Range row in doc.Body.Rows)
                {
                    if (++iRow == 1)
                    {
                        // занимаемся заголовками колонок -- пока просто пропустим строку
                        continue;
                    }
                    // цикл по проходам
                    foreach (pass passNum in Enum.GetValues(typeof(pass)))
                    {
                        int colNum = 0;         
                        foreach (Excel.Range col in doc.ptrn.Columns)
                        {
                            colNum++;           // ведем номер колонки в Range как целое число
                            string sX = col.Range[PTRN_COLS].Text;
                            string rqst = col.Range[PTRN_ADAPTER].Text;
                            int iX;
                            if (int.TryParse(sX, out iX))   // проверяем что число
                            {
                                if (passNum == pass.first) {
                                    // НЕДОПИСАНО!!! Надо извлечь номера отмеченных колонок из отдельного
                                    // именованного Range, например "HDR_1C_Payment_MyCol_Pass0"

                                    // на первом проходе - игнорируем все колонки кроме отмеченных
//                                    if (col.Range[PTRN_1STPASS].Text == "") continue;
                                    if (iX == colNum) continue;
                                } else if (iX != colNum) continue;
                                string x = row.Cells[1, 9].Text;
    //                            string x = row.Cells[1, iX].Text;
                                //                          string y = Adapter(rqst, 
                            }
                            else if (sX[0] == '#')
                            {
                                sX = sX.Substring(1);   // отсечь 1-й символ
                                if (int.TryParse(sX, out iX) || iX >= 0)   // проверяем что число и оно >= 0
                                {
                                }
                                else Log.FATAL("не числовое значение Шаблона с # в Value: '" + sX +"'");

                            }
                            else if (col.Range[PTRN_TITLE].Text == "ForProcess")
                            {
                            }
                            else Log.FATAL("недопустимое значение Шаблона в Value: '" + sX + "'");
                            // вызов адаптера rqst
                            string y = Adapter (rqst
                                            ,col.Range[PTRN_FETCH].Text
                                            );
                            if (y == null)  // Adapter возвращает null при ошибке
                            {
                            }
                        }
                    } 
                }
            }
            finally
            {
                Log.exit();
            }
 PK */
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
        string Adapter(string rqst, string fetch_rqst)
        {
            return null;
        }
    }
}
