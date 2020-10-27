using System;
using System.Collections.Generic;
using KPartition.Data;

namespace KPartition
{
    class Program
    {
        static void Main(string[] args)
        {
            var values = new List<int> {1,2,3,4,5,6,7,8,9,10,11,12};
            var numBoxes = 3;
            
            var set = values.ConvertAll(n => new Article(n));
            var res = BoxPacker.Pack(set, numBoxes);
            
            res.ForEach(b => Console.WriteLine(b));
        }
    }
}