using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MvcExcelExtensions.Helpers;
using MvcExcelExtensionsSample.Models;

namespace MvcExcelExtensionsSample.Controllers
{
    public class HomeController : ExcelMasterController
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ExportToExcel()
        {
            List<Person> lst = new List<Person>();
            for (int c = 0; c < 50; c++)
            {
                lst.Add(new Person() {BirthDate = new DateTime(1990, 12, 18), IsMale = true, Name = "Ibrahim", Summary = "I'm passionate about technology specially Microsoft's technologies in software development, Adore software development and willing to be one of the worldwide noted developers, getting the most of my technical skills along with my studies in business field to develop a real business solutions.I'm a knowledge hunger have no end point for my learning path and trying to reach for the sky." });
            }

            return Excel(lst, "TEST REPORT","TEST SHEET");
        }

        public ActionResult Documentation()
        {
            return View();
        }
    }
}