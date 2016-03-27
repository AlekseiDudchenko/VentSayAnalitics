using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VentSayAnalitics
{
    class AnaliticsMaterial
    {
        public string MaterialName { get; set; }

        public int MaterialIndex { get; set; }

        public string Date { get; set; }

        public string DocumentNomber { get; set; }

        public string Provider { get; set; }
  
        public double Debit { get; set; }

        public string DebitString { get; set; }

        public double Credit { get; set; }

        public string CreditString { get; set; }

        public double Cost { get; set; } 

        public string CostString { get; set; }
        
        public double Price { get; set; }

        public  string PriceString { get; set; }

        public double Balance { get; set; }

        public int Count { get; set; }






    }
}
