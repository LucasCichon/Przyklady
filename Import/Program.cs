using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Linq;

namespace Import
{
    class Program
    {
        static void Main(string[] args)
        {
            OptimaCOM.O_Login("ADMIN", "", "Firma_Demo");
            Console.WriteLine("zalogowano");

            Rejestr rejestr = new Rejestr();
            string plikPath = "";
            ImportClass.ImportZPliku(rejestr, plikPath);

            OptimaCOM.O_Logout();

            #region
            string[] directories = Directory.GetDirectories(@"C:\Banki", ".", SearchOption.AllDirectories);
            //List<string> rejestry = new List<string> { "1572", "1607", "1624" };

            List<Rejestr> rejestry = new List<Rejestr>()
            {
                new Rejestr{Numer = "1572", Nazwa="PK Pożyczki PRO WO RFP III"},
                new Rejestr{Numer = "1607", Nazwa="PK Pożyczki PT-Wypłąty Jednostkowe"},
                new Rejestr{Numer = "1624", Nazwa="PK Pożyczki PT-Zwroty Jednostkowe"}
            }.AsQueryable().ToList();
            
            foreach(string d in directories)
            {
                if (d.Contains("."))
                {
                Console.WriteLine(d);
                }
            }

            string data;
            Console.WriteLine("Podaj Datę:");
            data = Console.ReadLine();


            foreach (string d in directories)
            {
                if (d.Contains(data))
                {
                    Console.WriteLine(d);
                }
            }

            //mechanizm przerabiania daty na ścieżke pliku
            Char[] charsToRemove = { '.' };
            string[] pth = data.Split(charsToRemove);
            string Path="";
            for (int i=0; i< pth.Length; i++)
            {
                Path = Path + pth[i];
            }

            Console.WriteLine("ścieżka: "+Path);

            foreach(string s in directories)
            {
                if (s.Contains(data))
                {
                    Console.WriteLine(s);
                }
            }

            string endPath="";

            Console.WriteLine(endPath);
            string[] paths = Directory.GetFiles(@"C:\Banki\1607\2020");
            string[] lines = File.ReadAllLines(@"C:\Banki\1572\2020\03.2020\5041_02032020");


            foreach (string path in paths)
            {
                if (File.Exists(path))
                {
                    // This path is a file
                    Console.WriteLine(path);
                }
                else if (Directory.Exists(path))
                {
                    // This path is a directory
                    Console.WriteLine(path);
                }
                else
                {
                    Console.WriteLine("{ 0} is not a valid file or directory.", path);
                }
            }
            Console.ReadLine();
            #endregion

        }
    }
}
