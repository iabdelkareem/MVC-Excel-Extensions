using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using OfficeOpenXml.Style;

namespace MvcExcelExtensions.Options
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnStyleAttribute : Attribute
    {
        public double Width { get; set; }

        public HorizontalAlign HorizontalAlignment { get; set; }

        internal ExcelHorizontalAlignment ExcelHorizontalAlignment
        {
            get
            {
                switch (HorizontalAlignment)
                {
                    case HorizontalAlign.Center: return ExcelHorizontalAlignment.Center;
                    case HorizontalAlign.Justify: return ExcelHorizontalAlignment.Justify;
                    case HorizontalAlign.Left: return ExcelHorizontalAlignment.Left;
                    case HorizontalAlign.Right: return ExcelHorizontalAlignment.Right;
                    default: return ExcelHorizontalAlignment.General;
                }
            }
        }

        public VerticalAlign VerticalAlignment { get; set; }
        
        internal ExcelVerticalAlignment ExcelVerticalAlignment
        {
            get
            {
                switch (VerticalAlignment)
                {
                    case VerticalAlign.Top: return ExcelVerticalAlignment.Top;
                    case VerticalAlign.Middle: return ExcelVerticalAlignment.Center;
                    case VerticalAlign.Bottom: return ExcelVerticalAlignment.Bottom;
                    default: return ExcelVerticalAlignment.Center;
                }
            }
        }

        public bool WordWrap { get; set; }

        public ExcelColumnStyleAttribute()
        {
            Width = -1;
            WordWrap = false;
            VerticalAlignment = VerticalAlign.Middle;
            HorizontalAlignment = HorizontalAlign.Center;           
        }
    }
}
