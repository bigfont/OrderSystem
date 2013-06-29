using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OrderSystem.Models {
    public class ExcelRow {
        public int ExcelRowID { get; set; }
        public string ItemName { get; set; }
        public double ItemPrice { get; set; }
    }
}