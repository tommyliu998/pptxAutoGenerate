using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.Util;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pptxAutoGenetate
{
    class pptxUtil
    {
        private static int sraIndex = 0;
        private static int srlIndex = 0;
        private static int linePieChartIndex = 0;


        public static void GeneratePPT(string pptxFilePath, ProjectItem projectItem)
        {
            if (File.Exists(pptxFilePath))
            {
                File.Delete(pptxFilePath);
            }
            //initial data from project item
            string projectName = GetProjectName(projectItem);
            Dictionary<string, decimal> areaData = ParseAreaData(projectItem);
            Dictionary<string, Dictionary<string, decimal>> areaLineRelationData = ParseAreaLineRelationData(projectItem);
            Dictionary<string, Dictionary<string, decimal>> lineStationRelationData = ParseLineStationRelationData(projectItem);

            using (Presentation presentation = new Presentation(Constant.templateFilePath))
            {

                for (int i = 0; i < presentation.Slides.Count; i++)
                {
                    ISlide slide = presentation.Slides[i];
                    List<string> slidePrefixTexts = getAllPrefixTextForSlide(slide);
                    int textCount = slidePrefixTexts.Count;
                    for(int textIndex = 0; textIndex< textCount; textIndex++)
                    {
                        string text = slidePrefixTexts[textIndex];
                        if (text.Equals(Constant.projectTemplatePrefix))
                        {
                            ReplaceTagForSlide(slide, text, projectName);
                            continue;
                        }
                        if (text.StartsWith(Constant.areaTemplatePrefix))
                        {
                            replaceLineOrStationTemplatePrefix(presentation, slide, areaLineRelationData, null, null, textCount, textIndex, Constant.areaTemplatePrefix);
                            continue;
                        }
                        if (text.StartsWith(Constant.lineTemplatePrefix))
                        {
                            List<string> normalTexts = getNormalTextForSlide(slide);
                            string areakey = normalTexts[1].Substring(3);
                            string areaSpecificIndex = normalTexts[1].Substring(0, 1);
                            replaceLineOrStationTemplatePrefix(presentation, slide, areaLineRelationData, areakey, areaSpecificIndex, textCount, textIndex, Constant.lineTemplatePrefix);
                            continue;

                        }
                        if (text.StartsWith(Constant.stationTemplatePrefix))
                        {
                            List<string> normalTexts = getNormalTextForSlide(slide);
                            string lineKey = normalTexts[0].Substring(4);
                            string lineSpecificIndex = normalTexts[0].Substring(0, 3);
                            replaceLineOrStationTemplatePrefix(presentation, slide, lineStationRelationData, lineKey, lineSpecificIndex, textCount, textIndex, Constant.stationTemplatePrefix);
                            continue;

                        }
                        if (text.Equals(Constant.pcaTemplatePrefix))
                        {
                            pieChart.generate3DPieChartForArea(presentation, slide, areaLineRelationData);
                            continue;

                        }
                        if (text.Equals(Constant.pclTemplatePrefix))
                        {
                            List<string> lineKeys = lineStationRelationData.Keys.ToList();
                            Dictionary<string, decimal> data = null;
                            if (lineStationRelationData.TryGetValue(lineKeys[linePieChartIndex], out data))
                            {
                                pieChart.generate3DPieChartForLine(presentation, slide, lineKeys[linePieChartIndex], data);
                            }
                            linePieChartIndex++;
                            continue;

                        }
                        if (text.Equals(Constant.aoTemplatePrefix))
                        {
                            replaceAreaOutlineTemplate(presentation, slide, areaLineRelationData);
                            continue;
                        }

                        if (text.Equals(Constant.sraTemplatePrefix))
                        {
                            List<string> areaKeys = areaData.Keys.ToList();
                            if (sraIndex < areaKeys.Count)
                            {
                                string areaSpefic = string.Format("{0}. {1}", sraIndex + 1, areaKeys[sraIndex]);
                                ReplaceTagForSlide(slide, Constant.sraTemplatePrefix, areaSpefic);
                                sraIndex++;
                            }
                            continue;

                        }

                        if (text.Equals(Constant.srlTemplatePrefix))
                        {
                            List<string> linesKeys = lineStationRelationData.Keys.ToList();
                            List<string> areaKeys = areaLineRelationData.Keys.ToList();
                            Dictionary<string, decimal> lines = null;
                            if (areaLineRelationData.TryGetValue(areaKeys[sraIndex - 1], out lines))
                            {
                                List<string> linesList = lines.Keys.ToList();
                                int lineIndex = linesList.FindIndex(item => item.Equals(linesKeys[srlIndex]));
                                string lineSpefic = string.Format("{0}.{1} {2}", sraIndex, lineIndex + 1, linesKeys[srlIndex]);
                                ReplaceTagForSlide(slide, Constant.srlTemplatePrefix, lineSpefic);
                                srlIndex++;
                            }
                            continue;
                        }
                    }
                    
                }
                presentation.Save(pptxFilePath, SaveFormat.Pptx);
            }
        }

        public static void ReplaceTagForSlide(ISlide slide, string strToFind, string strToReplace)
        {
            foreach (IShape shape in slide.Shapes)
            {
                if (shape is IAutoShape)
                {
                    IAutoShape ias = (IAutoShape)shape;
                    foreach (IParagraph ipa in ias.TextFrame.Paragraphs)
                    {
                         if (ipa.Text.Equals(strToFind))
                         {
                              ipa.Text = strToReplace;
                         } 
                    }
                }
            }
        }
        public static List<string> GetAllTemplatePrefix(Presentation presentation)
        {
            List<string> allTemplatePrefix = new List<string>();
            ITextFrame[] textFrame = SlideUtil.GetAllTextFrames(presentation, false);

            for (int i = 0; i < textFrame.Length; i++)
            {
                foreach (IParagraph ipa in textFrame[i].Paragraphs)
                {
                    foreach (IPortion iport in ipa.Portions)
                    {
                        if (iport.Text.StartsWith(Constant.templateStartPrefix) && iport.Text.EndsWith(Constant.templateEndPrefix))
                        {
                            if (!allTemplatePrefix.Contains(iport.Text))
                            {
                                allTemplatePrefix.Add(iport.Text);
                            }
                        }
                    }
                }
            }
            return allTemplatePrefix;
        }

        public static List<string> getAllPrefixTextForSlide(ISlide slide)
        {
            List<string> allTemplatePrefixForSlide = new List<string>();
            foreach (IShape shape in slide.Shapes)
            {
                if (shape is IAutoShape)
                {
                    IAutoShape ias = (IAutoShape)shape;
                    foreach (IParagraph ipa in ias.TextFrame.Paragraphs)
                    {
                       if (ipa.Text.StartsWith(Constant.templateStartPrefix) && ipa.Text.EndsWith(Constant.templateEndPrefix))
                       {
                           allTemplatePrefixForSlide.Add(ipa.Text);
                       }
                        
                    }
                }
            }
            return allTemplatePrefixForSlide;
        }

        public static List<string> getNormalTextForSlide(ISlide slide)
        {
            List<string> allNormalTextsForSlide = new List<string>();
            foreach (IShape shape in slide.Shapes)
            {
                if (shape is IAutoShape)
                {
                    IAutoShape ias = (IAutoShape)shape;
                    foreach (IParagraph ipa in ias.TextFrame.Paragraphs)
                    {
                        
                            if (ipa.Text.StartsWith(Constant.templateStartPrefix) && ipa.Text.EndsWith(Constant.templateEndPrefix))
                            {
                                continue;
                            }
                            if(!string.IsNullOrEmpty(ipa.Text))
                            {
                                allNormalTextsForSlide.Add(ipa.Text);
                            }
                            
                    }
                }
            }
            return allNormalTextsForSlide;
        }
        public static string GetProjectName(ProjectItem projectItem)
        {
            return projectItem.ProjectName;
        }
        public static Dictionary<string, decimal> ParseAreaData(ProjectItem projectItem)
        {
            Dictionary<string, decimal> data = new Dictionary<string, decimal>();
            foreach (AreaItem areaItem in projectItem.Areas)
            {
                data.Add(areaItem.ItemName, areaItem.ItemWeight);
            }
            return data;

        }

        public static Dictionary<string, Dictionary<string, decimal>> ParseAreaLineRelationData(ProjectItem projectItem)
        {
            Dictionary<string, Dictionary<string, decimal>> data = new Dictionary<string, Dictionary<string, decimal>>();
            foreach (AreaItem areaItem in projectItem.Areas)
            {
                Dictionary<string, decimal> linesDic = new Dictionary<string, decimal>();
                foreach (LineItem lineItem in areaItem.Lines)
                {
                    linesDic.Add(lineItem.ItemName, lineItem.ItemWeight);
                }
                data.Add(areaItem.ItemName, linesDic);
            }
            return data;

        }
        public static Dictionary<string, Dictionary<string, decimal>> ParseLineStationRelationData(ProjectItem projectItem)
        {
            Dictionary<string, Dictionary<string, decimal>> data = new Dictionary<string, Dictionary<string, decimal>>();
            foreach (AreaItem areaItem in projectItem.Areas)
            {
                foreach (LineItem lineItem in areaItem.Lines)
                {
                    Dictionary<string, decimal> stationDic = new Dictionary<string, decimal>();
                    foreach (StationItem stationItem in lineItem.Stations)
                    {
                        stationDic.Add(stationItem.ItemName, stationItem.ItemWeight);
                    }
                    data.Add(lineItem.ItemName, stationDic);
                }
            }
            return data;

        }

        private static void replaceLineOrStationTemplatePrefix(Presentation pres, ISlide slide, Dictionary<string, Dictionary<string, decimal>> data, string subjectKey, string specificIndex, int slideTextCount, int textIndex, string templatePrefix)
        {

            List<string> results = null;
            Dictionary<string, decimal> output = null;
            int resultCount = 0;
            int currentSlidePos = slide.SlideNumber - 1;
            if (subjectKey != null)
            {
                if (data.TryGetValue(subjectKey, out output))
                {
                    results = output.Keys.ToList();
                    resultCount = results.Count;
                }
            }
            else
            {
                results = data.Keys.ToList();
                resultCount = results.Count;
            }
            if (textIndex == 0)
            {
                int mod = resultCount % (slideTextCount - 2);
                int cy = resultCount / (slideTextCount - 2);
                int cycleCount = mod == 0 ? cy - 1 : cy;

                if (subjectKey == null)
                {
                    mod = resultCount % (slideTextCount - 1);
                    cy = resultCount / (slideTextCount - 1);
                    cycleCount = mod == 0 ? cy - 1 : cy;
                }
                //data count more than pptx slide text count,need create new slide
                for (int j = cycleCount; j >= 1; j--)
                {
                    ISlide newSlide = slide;

                    for (int k = 0; k < slideTextCount - 1; k++)
                    {
                        string result = string.Empty;
                        string resultSpefic = string.Empty;
                        if (j == cycleCount && mod != 0)
                        {
                            if (k + 1 < mod)
                            {
                                int elementIndex = j * (slideTextCount - 1) + k;
                                result = results.ElementAt(elementIndex);
                                resultSpefic = string.Format("{0} {1}", elementIndex + 1, result);
                                if (specificIndex != null)
                                {
                                    resultSpefic = string.Format("{0}.{1} {2}", specificIndex, elementIndex + 1, result);
                                }

                            }
                        }
                        else
                        {
                            int elementIndex = j * (slideTextCount - 1) + k;
                            result = results.ElementAt(elementIndex);
                            resultSpefic = string.Format("{0} {1}", elementIndex + 1, result);
                            if (specificIndex != null)
                            {
                                resultSpefic = string.Format("{0}.{1} {2}", specificIndex, elementIndex + 1, result);
                            }
                        }
                        int tagCloneIndex = k + 1;
                        ReplaceTagForSlide(slide, templatePrefix + tagCloneIndex + "}}", resultSpefic);
                    }
                    pres.Slides.InsertClone(currentSlidePos + j - 1, newSlide);
                }

            }

            int tagIndex = subjectKey == null ? textIndex : textIndex - 1;
            //if data count less than pptx slide prefix tag,replace it to empty
            if (tagIndex >= resultCount)
            {
                tagIndex += 1;
                ReplaceTagForSlide(slide, templatePrefix + tagIndex + "}}", "");

            }
            else
            {
                string result = results.ElementAt(tagIndex);
                string resultSpefic = string.Format("{0} {1}", textIndex, result);
                if (subjectKey == null)
                {
                    tagIndex += 1;
                    result = results.ElementAt(textIndex);
                    resultSpefic = string.Format("{0} {1}", tagIndex, result);
                    ReplaceTagForSlide(slide, templatePrefix + tagIndex + "}}", resultSpefic);
                }
                else
                {
                    if (specificIndex != null)
                    {
                        resultSpefic = string.Format("{0}.{1} {2}", specificIndex, textIndex, result);
                    }
                    ReplaceTagForSlide(slide, templatePrefix + textIndex + "}}", resultSpefic);
                }

            }
        }

        private static void replaceAreaOutlineTemplate(Presentation pres, ISlide slide, Dictionary<string, Dictionary<string, decimal>> data)
        {
            ISlide lineSpecificSlide = null;
            ISlide outLineCloneSlide = null;
            ISlide areaSpecificSlide = null;
            ISlide linePieChartSlide = null;

            int lineIncreaseIndex = 0;
            List<string> areaKeys = data.Keys.ToList();
            for (int ouIndex = 0; ouIndex < areaKeys.Count; ouIndex++)
            {
                Dictionary<string, decimal> preLineData = new Dictionary<string, decimal>();
                Dictionary<string, decimal> curLineData = new Dictionary<string, decimal>();
                int preSlideLineCount = 0;
                int curSlideLineCount = 0;
                if (ouIndex - 1 >= 0)
                {
                    if (data.TryGetValue(areaKeys[ouIndex - 1], out preLineData))
                    {
                        preSlideLineCount = preLineData.Count;
                    }
                }
                if (data.TryGetValue(areaKeys[ouIndex], out curLineData))
                {
                    curSlideLineCount = curLineData.Count;
                }
                int currentIndex = slide.SlideNumber - 1;
                if (preSlideLineCount == 0)
                {
                    lineSpecificSlide = pres.Slides[currentIndex + 2];
                    linePieChartSlide = pres.Slides[currentIndex + 3];

                    for (int k = 1; k < curSlideLineCount; k++)
                    {
                        pres.Slides.InsertClone(currentIndex + 2 * k + 1, linePieChartSlide);
                        pres.Slides.InsertClone(currentIndex + 2 * k + 2, lineSpecificSlide);
                    }
                    lineIncreaseIndex += 2 * curSlideLineCount + 2;
                }
                else
                {

                    outLineCloneSlide = pres.Slides.AddClone(slide);
                    areaSpecificSlide = pres.Slides[currentIndex + 1];
                    lineSpecificSlide = pres.Slides[currentIndex + 2];
                    linePieChartSlide = pres.Slides[currentIndex + lineIncreaseIndex - 1];

                    string tmpAreaOutlineSpefic = string.Format("{0}. {1}", ouIndex + 1, areaKeys[ouIndex]);
                    ReplaceTagForSlide(outLineCloneSlide, Constant.bodyTextForAreaTemplatePrefix, "");
                    ReplaceTagForSlide(outLineCloneSlide, Constant.aoTemplatePrefix, tmpAreaOutlineSpefic);

                    pres.Slides.InsertClone(currentIndex + lineIncreaseIndex, outLineCloneSlide);
                    pres.Slides.InsertClone(currentIndex + lineIncreaseIndex + 1, areaSpecificSlide);
                    outLineCloneSlide.Remove();
                    for (int k = 1; k <=curSlideLineCount; k++)
                    {
                        int baseIndex = currentIndex + lineIncreaseIndex + 2 * k;
                        pres.Slides.InsertClone(baseIndex, lineSpecificSlide);
                        pres.Slides.InsertClone(baseIndex + 1, linePieChartSlide);

                    }

                    lineIncreaseIndex += 2 * curSlideLineCount + 2;
                }

            }

            string areaOutlineSpefic = string.Format("{0}. {1}", 1, areaKeys[0]);
            ReplaceTagForSlide(slide, Constant.bodyTextForAreaTemplatePrefix, "");
            ReplaceTagForSlide(slide, Constant.aoTemplatePrefix, areaOutlineSpefic);
        }
    }
}
