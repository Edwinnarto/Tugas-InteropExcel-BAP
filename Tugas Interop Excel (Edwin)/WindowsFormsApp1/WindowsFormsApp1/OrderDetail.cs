using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class OrderDetail
    {
        public int Nomor { get; set; }

        public int KodeBarang { get; set; }
        
        public string NamaBarang { get; set; }

        public int Quantity { get; set; }

        public string Satuan { get; set; }
    }
}
