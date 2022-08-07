using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K1020_PackingList.Models
{
    public class Serials_model
    {
        public int Id { get; set; }
        public int MainBoxNr { get; set; }
        public string BoxNr { get; set; }
        public string SerialNr { get; set; }
        public string PaletId { get; set; }
        public string PaletName { get; set; }//+
        public string CodeId { get; set; }
        public string CodeName { get; set; }//+
        public string SimCard { get; set; }//+
        public string Version { get; set; }//+
        public string CountryCode { get; set; }//+
        public int BatteryCount { get; set; }//+
        public DateTime AddDateTime { get; set; }
        public DateTime ModDateTime { get; set; }
        public string UniqNr { get; set; }
    }
}
