using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointWriter.Events
{
    public class PowerPointTableInterceptedEventArgs : EventArgs
    {
        protected PowerPoint.Table mTable;
        public PowerPointTableInterceptedEventArgs(PowerPoint.Table table)
        {
            mTable = table;
        }

        public PowerPoint.Table Table
        {
            get { return mTable; }
        }
    }
}
