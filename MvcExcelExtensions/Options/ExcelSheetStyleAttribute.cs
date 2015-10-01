using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MvcExcelExtensions.Options
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelSheetStyleAttribute : Attribute
    {
        public float HeaderFontSize { get; set; }

        public float BodyFontSize { get; set; }
        
        public string FontFamily { get; set; }
       
        public bool IsHeaderBold { get; set; }

        public Color HeaderBackgroundColor { get; set; }

        public ExcelSheetStyleAttribute()
        {
            HeaderFontSize = 14;
            BodyFontSize = 12;
            FontFamily = "Times New Roman";
            IsHeaderBold = true;
            HeaderBackgroundColor = Color.CadetBlue;
        }
    }
}
