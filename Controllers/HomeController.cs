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
            IEnumerable<ExcelRow> rows = GetAllExcelRows();
            return View(rows.ToList());
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file) {
            var filePath = UploadFile(file);
            var enumerable = GetExcelRowEnumerableFromExcelFile(filePath);
            UploadExcelRowEnumerableToDB(enumerable);
            return RedirectToAction("Index");
        }

        private IEnumerable<ExcelRow> GetAllExcelRows() {
            OrderSystem.DAL.OrderSystem db = new OrderSystem.DAL.OrderSystem();
            var excelRows = from d in db.ExcelRow
                            select d;
            return excelRows;
        }

        private void UploadExcelRowEnumerableToDB(IEnumerable<ExcelRow> enumerable) {
            OrderSystem.DAL.OrderSystem db = new OrderSystem.DAL.OrderSystem();
            foreach (ExcelRow row in enumerable) {
                db.ExcelRow.Add(row);
            }
            db.SaveChanges();
        }

        private IEnumerable<ExcelRow> GetExcelRowEnumerableFromExcelFile(string filePath) {
            var excel = new ExcelQueryFactory();
            excel.FileName = filePath;
            IEnumerable<ExcelRow> query = from e in excel.Worksheet<ExcelRow>(0)
                                          select e;
            return query;
        }

        private string UploadFile(HttpPostedFileBase file) {
            string fileName, filePath;
            filePath = null;
            if (file.ContentLength > 0) {
                fileName = Path.GetFileName(file.FileName);
                filePath = Path.Combine(Server.MapPath("~/App_Data/uploads"), fileName);
                file.SaveAs(filePath);
            }
            return filePath;
        }
    }
}
