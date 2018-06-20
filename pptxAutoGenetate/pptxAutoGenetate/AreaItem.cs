using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pptxAutoGenetate
{
    class AreaItem
    {
        public string ItemName { get; set; }
        public decimal ItemWeight { get; set; }
        public List<LineItem> Lines { get; set; }
    }
}
