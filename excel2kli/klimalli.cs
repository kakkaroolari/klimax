using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel2kli
{
    class klimalli
    {
        public int jarjestysnumero { get; set; }
        public DateTime aika { get; set; }
        public decimal sade { get; set; }
        public decimal sateily { get; set; }
        public decimal T_e { get; set; }
        public decimal RH_e { get; set; }
        public decimal T_i { get; set; }
        public decimal RH_i { get; set; }

        public override string ToString()
        {
            return $"{jarjestysnumero}\t{DCM(sade)}\t{DCM(sateily)}\t{DCM1(T_e)}\t{DCM(RH_e)}\t{DCM1(T_i)}\t{DCM(RH_i)}";
        }

       private string DCM1(decimal myvar)
       {
          //return String.Format("{0:0.##}", myvar, System.Globalization.CultureInfo.InvariantCulture);
          return myvar.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture);
       }

      private string DCM(decimal myvar)
        {
            //return String.Format("{0:0.##}", myvar, System.Globalization.CultureInfo.InvariantCulture);
            return myvar.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}
