using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Import
{
    class Program
    {
        static void Main(string[] args)
        {
            OptimaCOM.O_Login("ADMIN", "", "Firma_Demo");
            Console.WriteLine("zalogowano");
            
            ImportClass.ImportZPliku();

            OptimaCOM.O_Logout();
            //Console.ReadLine();
        }
    }
}
