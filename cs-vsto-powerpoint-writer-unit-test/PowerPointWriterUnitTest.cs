using System;
using Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointWriter;

namespace cs_vsto_powerpoint_writer_unit_test
{
    [TestClass]
    public class PowerPointWriterUnitTest
    {
        [TestMethod]
        public void TestUpdatePowerPoint()
        {
            PowerPointReportModifier builder = new PowerPointReportModifier();
            builder.ChartIntercepted += (sender, e) =>
            {
                string title = e.Title;
                PowerPoint.Chart chart = e.Chart;
                Worksheet sheet = e.Worksheet;

                // code to modify the chart here
            };
            builder.TableIntercepted += (sender, e) =>
            {
                PowerPoint.Table table = e.Table;

                // code to modify the table here
            };
            builder.TextFrameIntercepted += (sender, e) =>
            {
                PowerPoint.TextRange paragraph = e.Paragraph;
                
                // code to modify the paragraph here
            };

            builder.Apply("input.ppt", "output.ppt");
        }

    }
}
