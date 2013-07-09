using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OrderSystem.Models;
using LinqToExcel;
using System.Collections;

namespace OrderSystem.Controllers {
    public class HomeController : Controller {

        public ActionResult Index() {

            string excelFilePath, excelFileBaseName;
            excelFileBaseName = "Test1.xlsx";
            excelFilePath = Path.Combine(Server.MapPath("~/App_Data/uploads"), excelFileBaseName);
            var vendorItems = LinqToExcel_PropertyToColumnMapping(excelFilePath);

            return View(vendorItems);

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
        private IEnumerable<VendorItem> LinqToExcel_ByWorkSheetName(string excelFileFullName) {

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
