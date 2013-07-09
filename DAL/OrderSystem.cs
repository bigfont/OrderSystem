using System;
using System.Collections.Generic;
using System.Data.Entity;
using OrderSystem.Models;
using System.Data.Entity.ModelConfiguration.Conventions;

namespace OrderSystem.DAL {
    public class OrderSystem : DbContext {
        public DbSet<VendorItem> ExcelRow { get; set; }
    }
}