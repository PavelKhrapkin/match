/*-----------------------------------------------------------------------
 * Handler -- класс прорамм, участвующих в обработке доукментов проекта match 3.1
 * 
 *  30.01.2015  П.Храпкин, А.Пасс
 *  
 * -------------------------------------------
 * Handler(List<string> parameters, List<string> docNames)   - КОНСТРУКТОР заполняет каталог хендлеров
 * 
 * Шаг AccDeDup()   - Отчет о наличии Организаций - дубликатов в SFacc
 * Шаг ContDeDup()  - Отчет о наличии Контактов- дубликатов или однофамильцев в одной и той же Организации
 * 
 * !не отлажено! Шаг  NonDialogPass() - неинтерактивный процесс поиска новых Платежей для занесения в SF
  */
using System;
using System.Data;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Docs = match.Document.Document;
using match.Process;
using Log = match.Lib.Log;
using Decl = match.Declaration.Declaration;
using Lst = match.Document.Lst;
using FileOp = match.MyFile.FileOpenEvent;

namespace match.Handler
{
    class Handler
    {
        Dictionary<string,Docs> docs = new Dictionary<string, Docs>();
      
        public Handler(List<string> parameters, List<string> docNames)
        {
            foreach (string docName in docNames)
                if (!String.IsNullOrEmpty(docName)) docs.Add(docName, Docs.getDoc(docName));
        }

        /// <summary>
        /// Шаг AccDeDup()  - Отчет о наличии Организаций - дубликатов в SFacc
        /// </summary>
        /// <journal>24.01.2015</journal>
        public void AccDeDup()
        {
            Log.set("AccDeDup");
            Lst.Init(Lst.Entity.Accounts);
        //--- не понимаю толком как это устроено, но оно работает!!!
            var duplicates = Lst.Accounts.GroupBy(s => s).SelectMany(grp => grp.Skip(1));

            int nDub = duplicates.Count();
            DataTable dt = new DataTable();
            dt.Columns.Add("№"); dt.Columns.Add("Организация");
            dt.Rows.Add( string.Format("В SalesForce обнаружено {0} дубликатов Организаций", nDub) );
            if (nDub > 0)
            {
                int n = 0;
                foreach (var toMerge in duplicates)
                {
                    DataRow rw = dt.Rows.Add();
                    rw[0] = ++n;
                    rw[1] = toMerge;
                }
                dt.Rows.Add("Рекомендую их слить в SFDC: Организации-Объединение организаций, и по списку выше");
            }
            FileOp.WrReport("test", dt);
            Log.exit();
        }
        /// <summary>
        /// Шаг ContDeDup() - Отчет о наличии Контактов- дубликатов или однофамильцев в одной и той же Организации
        /// </summary>
        /// <journal>26.01.2015</journal>
        public void ContDeDup()
        {
            Log.set("ContDedup");            
            const string AccId = "AccId", Acc = "Организация", Family = "Фамилия", GivName = "Имя";
            Docs doc = Docs.getDoc("SFcont");
            DataTable dtCont = new DataTable();           
        //--- положим данные из Документа SFcont в структуру DataTable
            dtCont.Columns.Add(AccId);
            dtCont.Columns.Add(Acc);
            dtCont.Columns.Add(Family);
            dtCont.Columns.Add(GivName);
            for (int i = 2; i <= doc.Body.iEOL(); i++ )
            {
                DataRow rw = dtCont.Rows.Add();
                rw[AccId]   = doc.Body[i, Decl.SFCONT_ACCID];
                rw[Acc]     = doc.Body[i, Decl.SFCONT_ACCNAME];
                rw[Family]  = doc.Body[i, Decl.SFCONT_FAMILY];
                rw[GivName] = doc.Body[i, Decl.SFCONT_GIVNAME];
            }
        //--- сортируем по Организациям и по фамилиям Контактов
            DataView dv = dtCont.DefaultView;
            dv.Sort = AccId + " asc" + ", " + Family + " asc";
            DataTable sdt = dv.ToTable();
        //--- получаем список пар номеров строк - дубликатов в sdt
            List<int> indx = new List<int>();
            string xAcc = null, xFamily = null;
            char xGivName = '.';  // первая буква имени в предыдущей строке
            for (int i = 0; i < sdt.Rows.Count; i++)
            {
                DataRow rw = sdt.Rows[i];
                string id = rw[AccId].ToString(), fam = rw[Family].ToString(), givName = rw[GivName].ToString();
                char nam; 
                if (givName == "") nam = '?'; else nam = givName.First();
                if ((id == xAcc) && (fam == xFamily) && (nam == xGivName)) { indx.Add(i-1); indx.Add(i); }
                if (  id != xAcc    ) xAcc = id;
                if ( fam != xFamily ) xFamily = fam;
                if ( nam != xGivName) xGivName = nam;
            }
        //--- формируем отчет в dtRep
            DataTable dtRep = new DataTable();
            foreach (var v in sdt.Columns) dtRep.Columns.Add();
            dtRep.Rows.Add("Ниже список пар/групп контактов - дубликатов");
            int iPrev = -1, grp = 0;
            string prevFam = null;
            foreach (int i in indx)
            {
                DataRow rw = sdt.Rows[i];
                if (i != iPrev)
                {
                    string fam = rw[Family].ToString();
                    if ((string)fam != prevFam)
                    {
                        prevFam = fam;
                        dtRep.Rows.Add(string.Format(".. {0} .............", ++grp));
                    }
                    iPrev = i;                    
                    DataRow rwRep = dtRep.Rows.Add();
                    for (int v = 0; v < dtRep.Columns.Count; v++) rwRep[v] = rw[v];
                }
            }
            FileOp.WrReport("test", dtRep);
            Log.exit();
        }
        /// <summary>
        /// !не дописано!
        /// Шаг NonDialogPass() - неинтерактивный процесс поиска новых Платежей для занесения в SF
        ///                       обрабатывает отчеты по Платежам 1С и Документы из Lst
        /// прототипы в VB:
        ///     - PaidAnalitics: NonDialogPass()
        ///     - AdaptEngine:   WrNewSheet - запись рекорда в лист CSV
        /// </summary>
        /// <returns>int nPayment - количентво найденных новых Платежей</returns>
        /// <journal>17.1.2015 PKh
        /// 31.01.2015 - FetchInit - раскоментировал, разбираюсь, что делать с неоднозначностью ключей
        /// </journal>
        public int NonDialogPass()
        {
            Log.set("NonDialogPass");
            int nPayment = 0;
            // отладка -- потом убрать в UnitTest
            Docs docPay = Docs.getDoc("Платежи");
            docPay.FetchInit();

            Docs newPay = Docs.NewSheet("NewPayment");

            newPay.Body.AddRow();
            string[] hh = { "One", "Two", "Three" };
            newPay.Body.AddRow(hh);

            //            Docs newDogovor = Docs.NewSheet("NewContract");

            newPay.saveDoc();
            Log.exit();
            return nPayment;
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
/* PK            if ((string)docAcc.dt.Rows[1][1] == docAcc.ptrn.String(1,1))
                            Log.FATAL("Попытка обработать уже обработанный Документ");

                        if (docAcc.Body.Range["A1"].Text == docAcc.ptrn.Range["A1"].Text)
                            Log.FATAL("Попытка обработать уже обработанный Документ");
                        //---  вставляем колонки по числу MyCol        
                        docAcc.Sheet.Range["A1", docAcc.Body.Cells[1, docAcc.MyCol]].EntireColumn.Insert();
                        //--- устанавливает ширину колонки листа по значениям в строке Шаблона Width
                        new Log("после вставки колонок");
                        int i = 1;
                        foreach (Excel.Range col in docAcc.ptrn.Columns)
                        {
                            string s = col.Range[Decl.PTRN_WIDTH].Text;
                            if (s == Decl.PTRN_COPYHDR) col.Range["A1"].Copy(docAcc.Body.Cells[1, i]);
                            string[] ar = s.Split('/');
                            float W;
                            if (!float.TryParse(ar[0], out W)) Log.FATAL("ошибка в строке Width шаблона = \""
                                + s + "\" при обработке Документа " + docAcc.name); 
                            docAcc.Body.Columns[i++].ColumnWidth = W;
                        }
                        //--- копируем колонки MyCol от верха до Body.EOL
                        docAcc.ptrn.Range["A1", docAcc.ptrn.Cells[2, docAcc.MyCol]].Copy(docAcc.Body.Range["A1"]);
                        docAcc.Body.Range["A2", docAcc.Body.Cells[docAcc.Body.Rows.Count, docAcc.MyCol]].FillDown();
                        //--- если есть --> формируем пятку
                        if (docAcc.SummaryPtrn != null) docAcc.SummaryPtrn.Copy(docAcc.Summary.Range["A2"]);
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
/* PK
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
                Docs docAcc    = docs[DOC_ACC_SYNONIMS];
                docAcc.Reset();
                Excel.Range Bdy = docAcc.Body;

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
//                        Excel.Range rw = docAcc.AddLine();
                        docAcc.Body.Range["A" + rowNum].Value = str.Trim();
                        docAcc.Body.Range["B" + rowNum].Value = row.Range[SYN_VALUE_COL].Text;
                        rowNum++;
                    }
                }
            }
            finally { Log.exit(); }
 
        }
 */
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
                Docs docAcc = docs.First().Value;
                int iRow = 0;
                foreach (Excel.Range row in docAcc.Body.Rows)
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
                        foreach (Excel.Range col in docAcc.ptrn.Columns)
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
        string Adapter(string rqst, string fetch_rqst)
        {
            return null;
        }
    }
}
