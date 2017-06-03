using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace PowerPointWriter
{
    public class PowerPointReportBuilder
    {
        public void Build(string inputFileName, string outputFileName)
        {
            PowerPoint.Application ppApp = new PowerPoint.Application();

            ppApp.Visible = Office.MsoTriState.msoTrue;

            PowerPoint.Presentation ppt = ppApp.Presentations.Open(inputFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
            Excel.Application excel = null;

            for (int i = 0; i < ppt.Slides.Count; ++i)
            {
                int slideIndex = i + 1;
                PowerPoint.Slide slide = ppt.Slides[slideIndex];
                //Console.WriteLine(slide.Name);

                int shapeCount = slide.Shapes.Count;
                for (int j = 0; j < shapeCount; ++j)
                {
                    int shapeIndex = j + 1;
                    PowerPoint.Shape shape = slide.Shapes[shapeIndex];

                    if (shape.HasChart == Office.MsoTriState.msoTrue)
                    {
                        PowerPoint.Chart chart = shape.Chart;

                        PowerPoint.ChartData data = chart.ChartData;
                        data.Activate();
                        Excel.Workbook workbook = data.Workbook;
                        Excel.Worksheet worksheet = workbook.Sheets[1];

                        string title = string.Empty;
                        if (chart.HasTitle)
                        {
                            title = chart.ChartTitle.Caption;
                        }

                        Intercept(ppt, chart, worksheet, title);

                        if (excel == null)
                        {
                            excel = workbook.Application;
                        }
                    }

                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        var textFrame = shape.TextFrame;
                        var textRange = textFrame.TextRange;
                        var paragraphs = textRange.Paragraphs(-1, -1);
                        foreach (PowerPoint.TextRange paragraph in paragraphs)
                        {
                            
                            Intercept(ppt, textFrame, textRange, paragraph);
                        }
                    }

                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        PowerPoint.Table table = shape.Table;


                        Intercept(ppt, table);

                    }
                }

            }


            ppt.SaveAs(outputFileName, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoFalse);
            ppt.Close();


            ppApp.Quit();
            ppApp = null;

            if (excel != null)
            {
                excel.Quit();
                excel = null;
            }
        }

        public event EventHandler<PowerPointTextFrameInterceptedEventArgs> TextFrameIntercepted;
        private void Intercept(PowerPoint.Presentation ppt, PowerPoint.TextFrame textFrame, PowerPoint.TextRange textRange, PowerPoint.TextRange paragraph)
        {
            if (TextFrameIntercepted != null)
            {
                TextFrameIntercepted(ppt, new PowerPointTextFrameInterceptedEventArgs(textFrame, textRange, paragraph));
            }
        }

        public event EventHandler<PowerPointChartInterceptedEventArgs> ChartIntercepted;
        private void Intercept(PowerPoint.Presentation ppt, PowerPoint.Chart chart, Excel.Worksheet worksheet, string title)
        {
            if (ChartIntercepted != null)
            {
                ChartIntercepted(ppt, new PowerPointChartInterceptedEventArgs(chart, worksheet, title));
            }
        }

        public event EventHandler<PowerPointTableInterceptedEventArgs> TableIntercepted;
        private void Intercept(PowerPoint.Presentation ppt, PowerPoint.Table table)
        {
            if (TableIntercepted != null)
            {
                TableIntercepted(ppt, new PowerPointTableInterceptedEventArgs(table));
            }
        }

        
    }
}
