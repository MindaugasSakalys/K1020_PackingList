using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K1020_PackingList.Models
{
    public class Pallets_model
    {
        public int Id { get; set; }
        public string PalletName { get; set; }
        public bool DonePallet { get; set; }
        public DateTime AddDate { get; set; }
        public string PhotoPath { get; set; }
        public bool Foto { get; set; }
    }
}
