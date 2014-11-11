// Derived from: http://www.leniel.net/2009/07/creating-excel-spreadsheets-xls-xlsx-c.html
using NPOI.HSSF.UserModel;
using NpoiSample.Helpers;
using NpoiSample.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace NpoiSample.Controllers
{
    public class HomeController : Controller
    {
        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult DownloadXls()
        {
            var excelHelper = new ExcelHelper();
            var personRepository = new PersonRepository();
            var people = personRepository.Get();

            try
            {
                // Only export 3 columns
                var properties = excelHelper.GetProperties(typeof(Person), new string[] { "PersonId", "FirstName", "LastName" });

                var workbook = excelHelper.CreateXls<Person>(people, properties);

                var memoryStream = new MemoryStream();

                workbook.Write(memoryStream);

                return File(memoryStream.ToArray(), "application/vnd.ms-excel", "People.xls");
            }
            catch (Exception ex)
            {
                TempData["Message"] = "Oops! Something went wrong.";

                return RedirectToAction("Error");
            }
        }

        [HttpGet]
        public ActionResult DownloadXlsx()
        {
            var excelHelper = new ExcelHelper();
            var personRepository = new PersonRepository();
            var people = personRepository.Get();

            try
            {
                // Only export 3 columns
                var properties = excelHelper.GetProperties(typeof(Person), new string[] { "PersonId", "FirstName", "LastName" });

                var workbook = excelHelper.CreateXlsx<Person>(people, properties);

                var memoryStream = new MemoryStream();

                workbook.Write(memoryStream);

                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "People.xlsx");
            }
            catch (Exception ex)
            {
                TempData["Message"] = "Oops! Something went wrong.";

                return RedirectToAction("Error");
            }
        }

        [HttpPost]
        public ActionResult UploadFile(IEnumerable<HttpPostedFileBase> files)
        {
            if (files == null || !files.Any())
                return View("Index");

            var excelHelper = new ExcelHelper();

            var people = new List<Person>();

            try
            {
                var properties = excelHelper.GetProperties(typeof(Person), new[] { "PersonId", "FirstName", "LastName" });

                foreach (var file in files)
                {
                    if (file.ContentLength > 0)
                    {
                        var data = excelHelper.ReadData<Person>(file.InputStream, file.FileName, properties);

                        people.AddRange(data);
                    }
                }
            }
            catch (Exception ex)
            {
                TempData["Message"] = "Oops! Something went wrong.";

                return RedirectToAction("Error");
            }

            return View("Index", people);
        }

        public ActionResult Error()
        {
            return View();
        }
    }
}