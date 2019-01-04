using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {

            char a = '1';
            char b = '2';

            int c = 0;
            c = a + b;

            Console.WriteLine("Wynik to {0}", c);

            int z = int.Parse(a.ToString());
            int k = int.Parse(b.ToString()); ;
            int wynik = z + k;

            Console.WriteLine("Wynik drugi to {0}", wynik);

            Console.ReadKey();
        }
    }
}
