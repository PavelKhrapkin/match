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
    /// класс Fetch используется для извлечения значению поля из Документа по запросу вида
    ///     SF/2:18/0
    /// в строке Fetch Шаблона Документа. Здесь SF - имя Документа, 2 - колонка с ключем,
    /// 18 - значение, извлекаемое по ключу, а /0 - указание, что при неудачном поиске ошибку
    /// не выдавать.
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
