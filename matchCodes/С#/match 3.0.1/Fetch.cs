using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Docs = match.Document.Document;
using Log = match.Lib.Log;

namespace match.Fetch
{
    /// <summary>
    /// ����� Fetch ������������ ��� ���������� �������� ���� �� ��������� �� ������� ����
    ///     SF/2:18/0
    /// � ������ Fetch ������� ���������. ����� SF - ��� ���������, 2 - ������� � ������,
    /// 18 - ��������, ����������� �� �����, � /0 - ��������, ��� ��� ��������� ������ ������
    /// �� ��������.
    /// </summary>
    /// <journal>11.1.14 PKh</journal>
    public class Fetch
    {
        private static Dictionary<string, Fetch> Documents = new Dictionary<string, Fetch>();

        private string name;
        private bool isInitiated = false;
        private Docs doc;

        Dictionary<string, string> vals = new Dictionary<string, string>();
        
        Fetch(string fetch_rqst)
        {
            string[] ar_rqst = fetch_rqst.Split('/');
            Docs doc = Docs.getDoc(ar_rqst[0]);
 //           foreach ()
        }
       
    }
}
