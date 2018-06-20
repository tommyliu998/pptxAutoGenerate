using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pptxAutoGenetate
{
    class Program
    {
        static void Main(string[] args)
        {
          
            pptxUtil.GeneratePPT(@"c:\workspace\data.pptx", JsonUtil.getProjectItemFromJsonFile(@"c:\workspace\data.json"));
        }
    }
}
