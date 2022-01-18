using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FTIB1Core.Models.SAPObject
{
    public class RDR1
    {
        public int LineNum { get; set; }
        public string ItemCode { get; set; }
        public string Dscription { get; set; }
        public decimal? DiscPrcnt { get; set; }

        public int Index { get; set; }
    }
}
