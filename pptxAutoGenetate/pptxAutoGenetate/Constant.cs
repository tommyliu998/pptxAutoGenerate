using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pptxAutoGenetate
{
    class Constant
    {
        public static string templateFilePath = AppDomain.CurrentDomain.BaseDirectory + @"PT-Vorlage_Folienbibliothek_intern.pptx";
        public static string projectTemplatePrefix = "{{PROJECTNAME}}";
        public static string areaTemplatePrefix = "{{AREA";
        public static string pcaTemplatePrefix = "{{PIECHARTFORAREA}}";
        public static string pclTemplatePrefix = "{{PIECHARTFORLINE}}";
        public static string aoTemplatePrefix = "{{OUTLINEFORARE}}";
        public static string sraTemplatePrefix = "{{SPECIFICFORAREA}}";
        public static string lineTemplatePrefix = "{{LINE";
        public static string srlTemplatePrefix = "{{SPECIFICFORLINE}}";
        public static string stationTemplatePrefix = "{{STATION";
        public static string bodyTextForAreaTemplatePrefix = "{{BODYTEXTFORARE}}";
        public static string templateStartPrefix = "{{";
        public static string templateEndPrefix = "}}";

        public static int area3DPieChartCountShow = 2;
        public static int line3DPieChartCountShow = 1;
    }
}
