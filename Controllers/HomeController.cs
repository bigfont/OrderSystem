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

            string excelFileFullName, excelFileBaseName;
            IEnumerable<VendorItem> vendorItems;

            //create the full name of the excel file
            excelFileBaseName = "Test1.xlsx";
            excelFileFullName = Path.Combine(Server.MapPath("~/App_Data/uploads"), excelFileBaseName);

            //populate the viewbag with the worksheet names            
            IEnumerable<string> worksheets = LinqToExcel_GetWorksheetNames(excelFileFullName);
            ViewBag.WorksheetNames = new List<SelectListItem>();
            foreach (string w in worksheets) {
                ViewBag.WorksheetNames.Add(new SelectListItem { Text = w.ToString(), Value = "0" });
            }

            //populate the viewbag with column names
            IEnumerable<string> columns = LinqToExcel_GetColumnNames(excelFileFullName, "Sheet1");
            ViewBag.ColumnNames = new List<SelectListItem>();
            foreach (string c in columns) {
                ViewBag.ColumnNames.Add(new SelectListItem { Text = c.ToString(), Value = "0" });
            }

            //get the model to which we will bind the view
            vendorItems = LinqToExcel_DefaultQuery(excelFileFullName);

            return View(vendorItems);

        }

        private IEnumerable<String> LinqToExcel_GetColumnNames(string excelFileFullName, string worksheetName) {

            var excel = new ExcelQueryFactory(excelFileFullName);
            var columnNames = excel.GetColumnNames(worksheetName);
            return columnNames;

        }

        /// <summary>
        ///  Query Worksheet Names
        /// </summary>
        /// <remarks>
        /// The GetWorksheetNames() method can be used to retrieve the list of worksheet names in a spreadsheet.
        /// </remarks>
        /// <param name="excelFileFullName"></param>
        /// <returns></returns>
        private IEnumerable<String> LinqToExcel_GetWorksheetNames(string excelFileFullName) {

            var excel = new ExcelQueryFactory(excelFileFullName);
            var worksheetNames = excel.GetWorksheetNames();
            return worksheetNames;                        

        }

        /// <summary>
        ///  Query a worksheet with a header row
        /// </summary>
        /// <remarks>
        /// The default query expects the first row to be the header row containing column names that match the property names on the generic class being used. 
        /// It also expects the data to be in the worksheet named "Sheet1".
        /// </remarks>
        /// <param name="excelFileFullName"></param>
        /// <returns></returns>
        private IEnumerable<VendorItem> LinqToExcel_DefaultQuery(string excelFileFullName) {

            var excel = new ExcelQueryFactory(excelFileFullName);
            var vendorItems = from v in excel.Worksheet<VendorItem>()
                                   select v;
            return vendorItems;
        }

        /// <summary>
        ///  Query a specific worksheet by name
        /// </summary>
        /// <remarks>
        /// Data from the worksheet named "Sheet1" is queried by default. 
        /// To query a worksheet with a different name, pass the worksheet name in as an argument.
        /// </remarks>
        /// <param name="excelFileFullName"></param>
        /// <returns></returns>
        private IEnumerable<VendorItem> LinqToExcel_ByWorksheetName(string excelFileFullName) {

            var excel = new ExcelQueryFactory(excelFileFullName);
            var vendorItems = from v in excel.Worksheet<VendorItem>("Sale Items") //worksheet name = 'Sale Items'
                              select v;

            return vendorItems;
        }

        /// <summary>
        /// Property to column mapping
        /// </summary>
        /// <remarks>
        /// Column names from the worksheet can be mapped to specific property names on the class by using the AddMapping() method. 
        /// The property name can be passed in as a string or a compile time safe expression.
        /// </remarks>
        /// <param name="excelFileFullName"></param>
        /// <returns></returns>
        private IEnumerable<VendorItem> LinqToExcel_PropertyToColumnMapping(string excelFileFullName) {

            var excel = new ExcelQueryFactory(excelFileFullName);

            excel.AddMapping<VendorItem>(v => v.ItemName, "What We Call Our Items"); //pass mapping as compile time safe expression
            excel.AddMapping("ItemPrice", "The Prices We Charge"); //pass mapping as string

            var vendorItems = from v in excel.Worksheet<VendorItem>("Wierd Columns") //worksheet name = 'Wierd Columns'
                              select v;

            return vendorItems;
        }

    }
}
