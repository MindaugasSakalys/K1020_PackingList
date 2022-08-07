using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K1020_PackingList.Models
{
    public class Code_model
    {
        public int Id { get; set; }
        public string CodeName { get; set; }
        public string CountryCode { get; set; }
        public string CountryName { get; set; }
        public string Version { get; set; }
        public int BatteryCount { get; set; }
        public bool Disabled { get; set; }
    }
}
