using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointWriter.Events
{
    public class PowerPointTextFrameInterceptedEventArgs : EventArgs
    {
        private PowerPoint.TextFrame mTextFrame;
        private PowerPoint.TextRange mTextRange;
        private PowerPoint.TextRange mParagraph;

        public PowerPointTextFrameInterceptedEventArgs(PowerPoint.TextFrame textFrame,PowerPoint.TextRange textRange, PowerPoint.TextRange paragraph)
        {
            this.mTextFrame = textFrame;
            this.mTextRange = textRange;
            this.mParagraph = paragraph;
        }

        public PowerPoint.TextFrame TextFrame
        {
            get { return mTextFrame; }
        }

        public PowerPoint.TextRange TextRange
        {
            get { return mTextRange; }
        }

        public PowerPoint.TextRange Paragraph
        {
            get { return mParagraph; }
        }
    }
}
