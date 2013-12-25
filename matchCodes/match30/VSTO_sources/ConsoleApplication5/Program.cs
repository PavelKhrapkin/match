using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Decl = match.Declaration.Declaration;
using Document = match.Document.Document;
using Log = match.Lib.Log;

namespace match30
{
    class Program
    {
        static void Main(string[] args)
        {
            DateTime _t = DateTime.Now; 
            Console.WriteLine("{0}> match 3.0.0.0 -- отладка", _t);

            Log.set("Program");
            new Log("getDoc(SF)");

            Document.getDoc("SF");
            Console.ReadLine();
        }
    }
}