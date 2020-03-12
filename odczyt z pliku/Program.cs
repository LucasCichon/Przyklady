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
            string[] lines = File.ReadAllLines(@"C:\Banki\1572\2020\03.2020\5041_02032020");
            int i = 0;
            foreach(string line in lines)
            {
                if (line.Contains("4:")) { i++; }
            }

            //tworzymy tablice obiektów
            Zapis[] zapisy = new Zapis[i];

            Console.WriteLine("Ilość lini: "+i);
            int s = 0;
            for(int j=i; j>0; j--)
            {
                if (lines[s].Contains("PL"))
                {

                }
            }

            Console.ReadLine();
        }
    }
}
