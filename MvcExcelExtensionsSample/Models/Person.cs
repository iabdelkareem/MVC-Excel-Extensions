using System;
using System.Drawing;
using System.Web.UI.WebControls;
using MvcExcelExtensions.Options;

namespace MvcExcelExtensionsSample.Models
{
    [ExcelSheetStyle(HeaderFontSize = 14, BodyFontSize = 12, FontFamily = "Times New Roman", IsHeaderBold = true)]
    public class Person
    {
        public string Name { get; set; }

        [ExcelValueFormat("dd, MMMM yyyy"), ExcelDisplayName("Birthday")]
        public DateTime BirthDate { get; set; }

        [ExcelIgnore]
        public bool IsMale { get; set; }

        [ExcelColumnStyle(HorizontalAlignment = HorizontalAlign.Left, VerticalAlignment = VerticalAlign.Top, Width = 35, WordWrap = true)]
        public string Summary { get; set; }
    }
}