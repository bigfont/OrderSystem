using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OrderSystem.Models {
    public class VendorItem {
        public int ItemID { get; set; }
        public string ItemName { get; set; }
        public double ItemPrice { get; set; }
    }
}