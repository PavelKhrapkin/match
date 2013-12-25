using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using match.Document;
using Decl = match.Declaration;


namespace match.MyFile
{
    class FileOpenEvent
    {
        private static Excel.Application _app = null;
/*
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
*/
        public static Excel.Workbook fileOpen(string name)
        {
            if( _app == null ) _app = new Excel.Application();
            _app.Visible = true;
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
                wb = _app.Workbooks.Open(Decl.Declaration.dirDBs + name);
                return wb;
            }
            catch (Exception ex)
            {

//--                System.Windows.Forms.MessageBox.Show("Ошибка> " + ex.Message
//--                    + "\n не открыт файл '" + dirDBs + name + "'");
                Console.WriteLine("CATCH FileOpen " + name);
                return null;
            }
        }
    }
}
