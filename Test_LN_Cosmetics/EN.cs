using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test_LN_Cosmetics
{
    public class EN //класс для описания электронной накладной
    {
        public int N { get; set; }

        public DateTime date { get; set; }
        public String orderNumber { get; set; }
        public DateTime orderDate { get; set; }
        public string INN { get; set; }
        public string KPP { get; set; }
        public string name { get; set; }
        public decimal ItogoSumma { get; set; }
        public List<Nomenclatura_string> list { get; set; }

        public EN()
        {
            list = new List<Nomenclatura_string>();
        }
    }

    public class Sertificat_string //сертификат
    {

        
        public int sert_vid { get; set; }
        public string sert_seriya { get; set; }
        public string sert_reg_n { get; set; }
        public int kol_seriya { get; set; }
        public DateTime date_vyd { get; set; }
        public string sert_kto_vydal { get; set; }
        public DateTime sert_srok { get; set; }
       
    }

    public class Nomenclatura_string //номернклатура
    {

        public string code { get; set; }
        public string nomenclatura { get; set; }
        public string nomenclaturaName { get; set; }
        public string proizvod { get; set; }
        public string strana { get; set; }
        public int kol { get; set; }
        public decimal cena { get; set; }
        public int nds { get; set; }
        public decimal jnvl_cena { get; set; }
        public decimal cena_bez_nds { get; set; }
        public string N_GTD { get; set; }
               
        public string shtrih_kod { get; set; }
        public DateTime date_reg { get; set; }
        public decimal jnvl_reestr_cena { get; set; }
        public decimal Itogo_string { get; set; }
        
        public List<Sertificat_string> sert_list { get; set; } //сертификаты
        public List<Seriya_string> seriya_list { get; set; } //серии
        public int jnvl { get; set; }
    }
   public class Seriya_string //серия
    {
  
       public string LP_seriya { get; set; }
       public DateTime datePrep { get; set; }
       public DateTime srok { get; set; }
       public int seriya_kol { get; set; }
       
    }

    
}
