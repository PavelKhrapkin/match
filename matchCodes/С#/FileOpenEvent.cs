using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Box = System.Windows.Forms.MessageBox;
using Lib = match.MatchLib;

namespace ExcelAddIn2
{
    class FileOpenEvent
    {
        private static Excel.Application _app;
        private const string dirDBs = "C:\\Users\\Pavel_Khrapkin\\Documents\\Pavel\\match\\matchDBs\\";    //временно!!!
        /// <summary>
        /// Open Workbook Event Handler
        /// </summary>
        /// <param name="Wb"></param>
        /// <journal>   15.12.2013
        /// </journal>
        public void newFile(Excel.Workbook Wb)
        {
            _app = Wb.Application;
            string name = Document.recognizeDoc(Wb);
            if (name == null) return;
            Document doc = Document.loadDoc(name, Wb);

            Box.Show("Файл '" + Wb.Name + "' распознан как " + name);   //почему-то вылетает!
         }

        public static Excel.Workbook fileOpen(string name)
        {
            foreach (Excel.Workbook W in _app.Workbooks)
            {
                if (W.Name == name)
                {
                    if (W.ActiveSheet.IsNullOrEmpty()) continue;
                    return W;
                }
            }
            Excel.Workbook wb;
            try
            {
                wb = _app.Workbooks.Open(dirDBs + name);
                return wb;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Ошибка> " + ex.Message
                    + "\n не открыт файл '" + dirDBs + name + "'");
                return null;
            }
        }

     }
}
