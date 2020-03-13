using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace odczyt_z_pliku
{
    class Zapis
    {
        private string Konto { get; set; }
        private DateTime Data { get; set; }
        private decimal Wartosc { get; set; }
        private string Symbol { get; set; }

        public Zapis() {
           
        }

        public Zapis(string Konto, DateTime Data, decimal Wartosc, string Symbol)
        {
            this.Konto = Konto;
            this.Data = Data;
            this.Wartosc = Wartosc;
            this.Symbol = Symbol;
        }
    }
}
