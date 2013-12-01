/*-----------------------------------------------------------------------
 * MatchLib -- библиотека общих подпрограмм проекта match 3.0
 * 
 *  1.12.2013  П.Храпкин, А.Пасс
 *  
 * - 20.11.13 переписано с VBA на С#
 * - 1.12.13 добавлен метод ToIntList
 * -------------------------------------------
 * EOL(Sh)                  - возвращает число непустых строк в листе Sh
 * ToIntList(s, separator)  - возвращает List<int>, разбирая строку s с разделителями separator
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace match
{
    public static class MatchLib {
        public static int EOL(Excel.Worksheet Sh)
        /*
         * возвращает число непустых строк листа Sh 
         *
         * 21.11.2013
         */
        {
            int i;
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
        /*
         *  преобразует строку, разделенную символами separator, в List<int>.
         *  не преобразуемые компоненты пропускаются.
         */
        public static List<int> ToIntList(string s, char separator)
        {
            string[] ar = s.Split(separator);
            List<int> ints = new List<int>();
            foreach (var item in ar) {
                int v;
                if (int.TryParse(item, out v)) ints.Add(v);
            }
            return ints;
        }
        
    }
}
