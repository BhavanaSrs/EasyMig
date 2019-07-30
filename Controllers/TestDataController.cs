using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using EasyMig;
using EasyMig.Utility;

namespace EasyMig.Controllers
{
    public class TestDataController : Controller
    {
        Utility.TestData utility = new Utility.TestData();

        // GET: TestData        
        public ActionResult Index(string txtRecords)
        {
            Int32 records = 0;
            List<Models.TestDataFile> data = null;

            if (!String.IsNullOrWhiteSpace(txtRecords))
                records = Convert.ToInt32(txtRecords);

            if (records > 0)
                utility.GenerateTestData(records);

            data = GetTestData();

            return View(data);
        }

        [HttpPost]
        public ActionResult GenerateTestData(FormCollection form)
        {
            string name = form["txtRecords"].ToString();
            return View("Index");
        }

        public ActionResult TestDataValidation()
        {
            return View();
        }

        //Plant
        //[HttpPost]
        public ActionResult GeneratMaterialPlantTestData1(string Records)
        {
            utility.GenerateMaterialPlantTestData(Convert.ToInt32(100));
            return View("Index");
        }

        public ActionResult GeneratMaterialSalesTestData1(string Records)
        {
            utility.GenerateMaterialSalesTestData(Convert.ToInt32(100));
            return View("Index");
        }

        public ActionResult GeneratMaterialClassificationData(string Records)
        {
            utility.GenerateMaterialClassificationTestData(Convert.ToInt32(100));
            return View("Index");
        }

        public ActionResult Validate()
        {
            var mandate = utility.ValidateExcelFile(@"C:\Projects\EasyMIG\BP sample data.xlsx");
            return View("TestDataValidation", mandate);
        }

        public List<EasyMig.Models.TestDataFile> GetTestData()
        {
            Utility.TestData utility = new Utility.TestData();
            return utility.GetTestData();
        }

        public ActionResult Download(string file)
        {
            if (!System.IO.File.Exists(file))
            {
                return HttpNotFound();
            }

            var fileBytes = System.IO.File.ReadAllBytes(file);
            var response = new FileContentResult(fileBytes, "application/octet-stream")
            {
                FileDownloadName = "file.xlsx"
            };
            return response;
        }

    }
}