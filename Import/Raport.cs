using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Import
{
    class Raport
    {
        public string NazwaRejestru { get; set; }
        public DateTime DataOtwarcia { get; set; }
        public DateTime DataZamkniecia { get; set; }
        public string DataOtwarciaString { get; set; }
        public string DataZamknieciaString { get; set; }
        public int NumerNr { get; set; }
    }
}
