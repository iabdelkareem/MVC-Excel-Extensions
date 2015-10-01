using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace MvcExcelExtensions.Helpers
{

    public class ExcelMasterController : Controller
    {

        protected ExcelResult<T> Excel<T>(IEnumerable<T> data) where T : class
        {
            return new ExcelResult<T>(data);
        }

        protected ExcelResult<T> Excel<T>(IEnumerable<T> data, string fileName) where T : class
        {
            return new ExcelResult<T>(data, fileName);
        }

        protected ExcelResult<T> Excel<T>(IEnumerable<T> data, string fileName, string sheetName) where T : class
        {
            return new ExcelResult<T>(data, fileName, sheetName);
        }

   }

}
