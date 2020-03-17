using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace odczyt_z_pliku
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"‪C:\Banki\1572\2020\03.2020\1572_02032020";
            if (File.Exists(path))
            {
                Console.WriteLine("Istnieje");
            }
            else
            {
                Console.WriteLine("Nie Istnieje");
            }

            Console.ReadLine();
        }
    }
}
