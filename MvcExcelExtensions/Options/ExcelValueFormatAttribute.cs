using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MvcExcelExtensions.Options
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelValueFormatAttribute : Attribute
    {
        public string Format { get; set; }

        public ExcelValueFormatAttribute(string format)
        {
            Format = format;
        }
    }
}
