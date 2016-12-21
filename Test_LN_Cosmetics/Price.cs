using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test_LN_Cosmetics
{
   public class Price
    {
       public int Cod { get; set; }
       public string Name { get; set; }
       public string MNF { get; set; }
       public string CNTR { get; set; }
       public DateTime Srok { get; set; }
       public int Kol { get; set; }
       public decimal Cena { get; set; }
       public int  Kratnost { get; set; }
       public Int64 Barcode { get; set; }
       public int Ratends { get; set; }

       public Price()
       {
           this.Cod = 0;
           this.Name = "";
           this.MNF = "";
           this.CNTR = "";
           this.Srok = DateTime.Parse("01/01/0001");
           this.Kol = 0;
           this.Cena = 0;
           this.Kratnost = 1;
           this.Barcode = 0;
           this.Ratends = 18;
       }
    }
}
