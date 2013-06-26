using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OrderSystem.Models;
using LinqToExcel;

namespace OrderSystem.Controllers {
    public class HomeController : Controller {
        public ActionResult Index() {
            ViewBag.Message = "Hello world.";
            var data = ReadExcel();
            return View(data);
        }

        /*
         * Upload a file.
         */
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file) {

            if (file.ContentLength > 0) {
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/App_Data/uploads"), fileName);
                file.SaveAs(path);
            }

            return RedirectToAction("Index");
        }

        private IEnumerable<ExcelRow> ReadExcel() {

            var fileName = Server.MapPath("~/App_Data/uploads/Test1.xlsx");
            var excel = new ExcelQueryFactory();
            excel.FileName = fileName;
            IEnumerable<ExcelRow> query = from e in excel.Worksheet<ExcelRow>(0)
                                          select e;

            return query;
        }
    }
}
