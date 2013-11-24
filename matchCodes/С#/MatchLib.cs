/*-----------------------------------------------------------------------
 * MatchLib -- библиотека общих подпрограмм проекта match 3.0
 * 
 *  20.11.2013  П.Храпкин, А.Пасс
 *  
 * - 20.11.13 переписано с VBA на С#
 * -------------------------------------------
 * EOL(Sh)          - возвращает число непустых строк в листе Sh
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace match
{
    class MatchLib {
        public static int EOL(Excel.Worksheet Sh)
        /*
         * возвращает число непустых строк листа Sh 
         *
         * 21.11.2013
         */
        {
            int i;
//            int j = Sh.UsedRange.Rows.Count;
//            System.Windows.Forms.MessageBox.Show("Sh.Name='"+Sh.Name+"'  Count="+j);
//            for (i = j; i > 0; i--) {
            for (i = Sh.UsedRange.Rows.Count; i > 0; i--) {
                for (int col = 1; col <= Sh.UsedRange.Columns.Count; col++)
                {
                    if (Sh.Cells[i, col].Value2 != null
                        /*                        || ((String)Sh.Cells[i, col].Value2).Trim() != "" */ )
                    {
                        return i;
                    }
                };
            }
            return i;
        }
        
    }
}
