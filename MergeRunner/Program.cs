using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeRunner
{
    class Program
    {
        static void Main(string[] args)
        {
            var merger = new WordMailMerge.Main();
            merger.DoMerge();
        }
    }
}
