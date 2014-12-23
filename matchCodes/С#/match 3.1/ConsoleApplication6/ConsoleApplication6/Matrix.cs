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
        private object[,] _matr = new object[100, 100];
        
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
        public int iEOL() { return _matr.GetLength(0); }
        public int iEOC() { return _matr.GetLength(1); }
        public DataTable DaTab()
        {
            DataTable _dt = new DataTable();
            int maxCol = iEOC(), maxRow = iEOL();
            for (int col = 1; col <= maxCol; col++) _dt.Columns.Add();
            for (int rw = 1; rw <= maxRow; rw++)
            {
                _dt.Rows.Add();
                for (int col = 1; col <= maxCol; col++) _dt.Rows[rw - 1][col - 1] = get(rw, col);
            }
            return _dt;
        }

    } // конец класса Matr
}
