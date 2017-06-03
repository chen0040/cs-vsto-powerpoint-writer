using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;

namespace PowerPointWriter.Events
{
    public class PowerPointChartInterceptedEventArgs : EventArgs
    {
        private string mTitle;
        private PowerPoint.Chart mChart;
        private Excel.Worksheet mWorksheet;

        public PowerPointChartInterceptedEventArgs(PowerPoint.Chart chart, Excel.Worksheet worksheet, string title)
        {
            this.mTitle = title;
            this.mChart = chart;
            this.mWorksheet = worksheet;
        }

        public string Title
        {
            get { return mTitle; }
        }

        public PowerPoint.Chart Chart
        {
            get { return mChart; }
        }

        public Excel.Worksheet Worksheet
        {
            get { return mWorksheet; }
        }
    }
}
