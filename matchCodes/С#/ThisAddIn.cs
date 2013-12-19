using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn2
{
    public partial class ThisAddIn
    {
 
        private FileOpenEvent check;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
//            System.Windows.Forms.MessageBox.Show("Hello");
//            this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            check = new FileOpenEvent();
            this.Application.WorkbookOpen += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookOpenEventHandler(Application_myOpenEvent);

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        void Application_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
 /*           Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Excel.Range firstRow = activeWorksheet.get_Range("A1");
            firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
            newFirstRow.Value2 = "This text was added by using code";
  */
            System.Windows.Forms.MessageBox.Show("Helolo 3");
        }

        void Application_myOpenEvent(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            if (Wb.Name == Document.F_MATCH
                || Wb.Name == Document.F_1C
                || Wb.Name == Document.F_SFDC
                || Wb.Name == Document.F_TMP
                || Wb.Name == Document.F_ADSK
                || Wb.Name == Document.F_STOCK
                ) return;
//            System.Windows.Forms.MessageBox.Show(Wb.Name);
            try {
                check.newFile(Wb);
            } catch (Exception ex) {
                System.Windows.Forms.MessageBox.Show("Ошибка> " + ex.Message);
            }
        }

        #endregion
    }
}
