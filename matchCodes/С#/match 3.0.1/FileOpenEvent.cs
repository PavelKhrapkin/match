using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using match.Document;
using match.Lib;
using Decl = match.Declaration.Declaration;

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
        /// <summary>
        /// открываем файл Excel по имени name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Excel.Workbook</returns>
        /// <journal>11.12.2013
        /// 6.1.14  - единая точка выхода из метода
        /// </journal>
        public static Excel.Workbook fileOpen(string name)
        {
            Log.set("fileOpen");
            if (_app == null) _app = new Excel.Application();
            _app.Visible = true;
            Excel.Workbook _wb;
            foreach (Excel.Workbook W in _app.Workbooks)
                if (W.Name == name) _wb = W;
            try { _wb = _app.Workbooks.Open(Decl.dirDBs + name); }
            catch (Exception ex)
            {
                
                Log.FATAL("не открыт файл" + name + "\nDirDBs=\"" + Decl.dirDBs + "\""
                    + "\n сообщение по CATCH '" + ex);
                return null;
            }
            Log.exit();
            return _wb;
        }
        public static void fileSave(Excel.Workbook Wb) { Wb.Save(); }
    }
}
