using Aspose.Slides;
using Aspose.Slides.Charts;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace pptxAutoGenetate
{
    class pieChart
    {
        public static void add3DPieChartToSlide(ISlide slide, ChartProperty chartProperty, Dictionary<string, decimal> data)
        {
            string title = chartProperty.title;
            float x = chartProperty.x;
            float y = chartProperty.y;
            float height = chartProperty.height;
            float width = chartProperty.width;

            IChart chart = slide.Shapes.AddChart(ChartType.Pie3D, x, y, width, height);

            // Setting chart Title
            chart.ChartTitle.AddTextFrameForOverriding(title);
            chart.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.True;
            chart.ChartTitle.TextFrameForOverriding.TextFrameFormat.WrapText = NullableBool.False;
            chart.ChartTitle.Height = 30;
            chart.HasTitle = true;

            int defaultWorksheetIndex = 0;

            int index = 0;
            // Getting the chart data worksheet
            IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;

            // Delete default generated series and categories
            chart.ChartData.Series.Clear();
            chart.ChartData.Categories.Clear();


            // Adding new series
            IChartSeries series = chart.ChartData.Series.Add(fact.GetCell(0, 0, 1, title), chart.Type);
            chart.ChartData.SeriesGroups[0].IsColorVaried = true;
           
            var sortedData = from objDic in data orderby objDic.Value descending select objDic;
            var sum = sortedData.Select(item => item.Value).Sum();
            foreach (KeyValuePair<string, decimal> pair in sortedData)
            {
                // Adding new categories
                double percent = (double) (pair.Value / sum);
                string percentStr = percent.ToString("0.%");
                chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, index, 0, pair.Key+":"+ percentStr));
                // Now populating series data
                series.DataPoints.AddDataPointForPieSeries(fact.GetCell(defaultWorksheetIndex, index, 1, pair.Value));
                index++;
            }

            // Create custom labels for each of categories for new series
            for (int i = 0; i < series.DataPoints.Count; i++)
            {
                IDataLabel lbl = series.DataPoints[i].Label;
                lbl.TextFrameForOverriding.TextFrameFormat.WrapText = NullableBool.False;
                lbl.DataLabelFormat.Position = LegendDataLabelPosition.InsideEnd;
                lbl.DataLabelFormat.ShowCategoryName = false;
                lbl.DataLabelFormat.ShowPercentage = true;
                lbl.DataLabelFormat.ShowLeaderLines = false;
                lbl.DataLabelFormat.Separator = ",";
            }

            // no showing Legend key in chart
            series.Chart.HasLegend = true;
            // Set Legend Properties
            chart.Legend.Width = 200 / chart.Width;
            chart.Legend.Height = 100 / chart.Height;

            // Showing Leader Lines for Chart
            series.Labels.DefaultDataLabelFormat.ShowLeaderLines = false;
            series.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FillFormat.SolidFillColor.Color = Color.Black;
            series.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FillFormat.FillType = FillType.Solid;
            chart.Legend.Position = LegendPositionType.Bottom;
        }

        public static void generate3DPieChartForArea(Presentation pres, ISlide slide, Dictionary<string, Dictionary<string, decimal>> data)
        {
            int currentSlidePos = slide.SlideNumber - 1;
            int cy = data.Keys.Count / Constant.area3DPieChartCountShow;
            int mod = data.Keys.Count % Constant.area3DPieChartCountShow;
            int pieChartSlideCount = mod == 0 ? cy : cy + 1;

            List<string> dataKeys = data.Keys.ToList();

            ChartProperty cp = new ChartProperty();
            cp.y = 100;
            cp.height = 350;
            cp.width = 350;

            for (int j = 1; j <= pieChartSlideCount; j++)
            {
                List<string> tmpDataKeys = new List<string>();
                Dictionary<string, decimal> dataSpecific = null;
                List<Dictionary<string, decimal>> tempDataSpecific = new List<Dictionary<string, decimal>>();
                tmpDataKeys.Add(dataKeys[j * 2 - 2]);
                if (data.TryGetValue(dataKeys[j * 2 - 2], out dataSpecific))
                {
                    tempDataSpecific.Add(dataSpecific);
                }
                if (j * 2 - 1 < dataKeys.Count)
                {
                    tmpDataKeys.Add(dataKeys[j * 2 - 1]);
                    if (data.TryGetValue(dataKeys[j * 2 - 1], out dataSpecific))
                    {
                        tempDataSpecific.Add(dataSpecific);
                    }
                }
               
                if (j != 1)
                {
                    ISlide CloneSlide = pres.Slides.AddClone(slide);
                    pptxUtil.ReplaceTagForSlide(CloneSlide, Constant.pcaTemplatePrefix, "");
                    for (int index = 0; index < tempDataSpecific.Count; index++)
                    {
                        cp.x = 50 + index * cp.width + index * 20;
                        cp.title = string.Format("{0} progress", tmpDataKeys[index]);
                        add3DPieChartToSlide(CloneSlide, cp, tempDataSpecific[index]);
                    }
                    if (pres.Slides.Count - 1 > currentSlidePos + j)
                    {
                        pres.Slides.InsertClone(currentSlidePos + j -1, CloneSlide);
                    }
                    CloneSlide.Remove();
                }
                else
                {
                    ISlide CloneSlide = pres.Slides.AddClone(slide);
                    pptxUtil.ReplaceTagForSlide(CloneSlide, Constant.pcaTemplatePrefix, "");
                    for (int index = 0; index < tempDataSpecific.Count; index++)
                    {
                        cp.x = 50 + index * cp.width + index * 20;
                        cp.title = string.Format("{0} progress", tmpDataKeys[index]);
                        add3DPieChartToSlide(CloneSlide, cp, tempDataSpecific[index]);
                    }
                    pres.Slides.InsertClone(currentSlidePos, CloneSlide);
                    CloneSlide.Remove();
                }

            }

            slide.Remove();
        }

        public static void generate3DPieChartForLine(Presentation pres, ISlide slide, string lineKey, Dictionary<string, decimal> data)
        {
            
            ChartProperty cp = new ChartProperty();
            cp.x = 200;
            cp.y = 100;
            cp.height = 350;
            cp.width = 400;
            cp.title = string.Format("{0} progress", lineKey);

            ISlide currentSlide = slide;
            pptxUtil.ReplaceTagForSlide(currentSlide, Constant.pclTemplatePrefix, "");
            add3DPieChartToSlide(currentSlide, cp, data);
        }
    }
}
