using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MvcExcelExtensions.Options
{
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ExcelDisplayNameAttribute : Attribute
    {
        public string Name { get; set; }

        public ExcelDisplayNameAttribute(string name)
        {
            Name = name;
        }
    }
}
