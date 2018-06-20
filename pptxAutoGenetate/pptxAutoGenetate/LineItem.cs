using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pptxAutoGenetate
{
    class LineItem
    {
        public string ItemName { get; set; }
        public decimal ItemWeight { get; set; }
        public List<StationItem> Stations { get; set; }
    }
}
