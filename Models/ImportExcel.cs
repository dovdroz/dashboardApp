using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace DashApp.Models
{
    public class ImportExcel
    {
        public string Key { get; set; }
        public string KeyLink { get; set; }
        public string Summary { get; set; }
        public string Status { get; set; }
        public DateTime Created { get; set; }
        public DateTime Resolved { get; set; }
    }
}