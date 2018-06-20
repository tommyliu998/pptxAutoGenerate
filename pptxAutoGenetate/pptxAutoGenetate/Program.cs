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
            var pptPath = AppDomain.CurrentDomain.BaseDirectory + "data.pptx";
            var json = AppDomain.CurrentDomain.BaseDirectory + "data.json";
            pptxUtil.GeneratePPT(pptPath, JsonUtil.getProjectItemFromJsonFile(json));
        }
    }
}
