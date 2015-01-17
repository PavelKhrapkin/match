using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Log = match.Lib.Log;

namespace match.Matrix
{
    //////public class Matrix
    //////{
    //////    object value;

    //////    public Matrix(object val)
    //////    {
    //////        value = val;
    //////    }

    //////    public object get() {return value;}
    //////    public string ToStr() { return (value == null) ? "" : value.ToString(); }
    //////    public int ToInt(string msg)
    //////    {
    //////        if (value == null) return 0;
    //////        try
    //////        {
    //////            if (value.GetType() == typeof(int)) { return (int)value; }
    //////            int v;
    //////            if (int.TryParse(value.ToString(), out v)) return v;
    //////            Log.FATAL(msg);
    //////        }
    //////        catch { Log.FATAL(msg); }
    //////        return 0;
    //////    }
    //////}

    public class Matr : Object
    {
        private const int MATR_PAGE = 100;
        private object[,] _matr = new object[MATR_PAGE, MATR_PAGE];
        ////private int _EOL = 0;
        
        public Matr(object[,] obj)
        {
            _matr = obj;
        }
        public Matr(DataTable dt)
        {
            try
            {
                foreach (DataRow row in dt.Rows)
                {
                    int rw = 0;
                    object obj;
                    for (int col = 0; col <= dt.Columns.Count; col++)
                    {
                        obj = row[col];
                        _matr[rw++, col] = obj;
                    }
                }
            }
            catch (Exception ex)
            {
                string mes = ex.Message;
            }
        }
#if пока_не_нужно
        public List<object> getRow(int iRow)
        {
            List<object> _row = new List<object>(); 
            for (int i = 0; i < iEOC(); i++) _row.Add(_matr[iRow, i]);
            return _row;
        }
        public List<object> getCol(int col)
        {
            List<object> _col = new List<object>(); 
            for (int i = 0; i < iEOC(); i++) _row.Add(_matr[iRow, i]);
            return _row;
        }
#endif
        public object get(int i, int j)
        {
            object v = null;
            try { v = _matr[i, j]; }
            catch { Log.FATAL("ошибка при обращении к Matr[" + i + "," + j + "]"); }
            return v;
        }
        public void set(int i, int j)
        {
            try { _matr[i, j] = this; }
            catch { Log.FATAL("!"); }
        }
        /////// <summary>
        /////// setRow(int i, object[] obj) - записывает ряд значений в Matr
        /////// </summary>
        /////// <param name="i">    int i   - номер рекорда матрицы</param>
        /////// <param name="object[] obj"> object[] obj - массив объектов, записываемых в матрицу</param>
        /////// <journal>28.12.2014
        /////// </journal>
        ////public void setRow(int i, object[] obj)
        ////{
        ////    try { for (int col = 1; col <= this.iEOC(); col++) { _matr[i, col] = obj[col]; } }
        ////    catch { Log.FATAL("! строка " + i); }
        ////}
        /////// <summary>
        /////// int AddRow(object[] obj) - добавляет ряд значений в Matr
        /////// </summary>
        /////// <param name="object[] obj">object[] obj - массив объектов, записываемых в матрицу</param>
        /////// <journal>28.12.2014
        /////// </journal>
        ////public void AddRow(object[] obj)
        ////{
        ////    try
        ////    {
        ////        if (_EOL++ > MATR_PAGE)
        ////        {
        ////            Log.FATAL("!!!! не написано");
        ////        }
        ////        setRow(_EOL, obj);
        ////    }
        ////    catch { Log.FATAL("! EOL=" + _EOL); }
        ////}
        public string String(int i, int j)
        {
            var v = get(i, j);
            return (v == null) ? "" : v.ToString();
        }
        public int Int(int i, int j, string msg = "wrong int")
        {
            object v = get(i, j);
            if (v == null) return 0;
            if (v.GetType() == typeof(int)) { return (int)v; }
            try
            {
                int value;
                string val = v.ToString();
                if (int.TryParse(val, out value)) return value;
                Log.FATAL(msg);
            }
            catch { Log.FATAL(msg); }
            return 0;
        }
        public float Float(int i, int j, string msg = "wrong Float")
        {
            object v = get(i, j);
            if (v == null) return 0;
            if (v.GetType() == typeof(float)) { return (float)v; }
            try
            {
                int value;
                string val = v.ToString();
                if (int.TryParse(val, out value)) return value;
                Log.FATAL(msg);
            }
            catch { Log.FATAL(msg); }
            return 0;
        }
        public int iEOL() { return _matr.GetLength(0); }
        public int iEOC() { return _matr.GetLength(1); }

        /// <summary>
        /// копируем данные из (matr)this в (Data Table)
        /// </summary>
        /// <returns></returns>
        /// <journal>2014
        /// 15.01.15 buf fix: все индексы dt и параметры get ведем в диапазоне 0..iEO, а не 1..iEO 
        /// </journal>
        public DataTable DaTab()
        {
            DataTable _dt = new DataTable();
            int maxCol = iEOC(), maxRow = iEOL();
            for (int col = 0; col < maxCol; col++) _dt.Columns.Add();
            for (int rw = 0; rw < maxRow; rw++)
            {
                _dt.Rows.Add();
    //!!            for (int col = 0; col < maxCol; col++) _dt.Rows[rw][col] = get(rw, col);
            }
            return _dt;
        }
        /// <summary>
        /// AddRow() - добавляет одну строку в конце матрицы
        /// </summary>
        /// <journal> 17.01.15 </journal>
        public Matr AddRow()
        {
            int irw = iEOL(), icol = iEOC();

            var newMatr = new object[irw + 1, icol];
            for (int i = 0; i < irw; i++)
                for (int j = 0; j < icol; j++)
                    newMatr[i,j] = get(i, j);
            for (int j = 0; j < icol; j++) newMatr[irw, j] = null;
            _matr = newMatr;
            return this;
        }
        public Matr AddRow(object[] Line)
        {
            AddRow();
            int rw = iEOL() -1;
            int cols = Math.Min(iEOC(), Line.Length);
            for (int i = 0; i < cols; i++) _matr[rw, i] = Line[i];
            return this;
        }
    } // конец класса Matr
}
