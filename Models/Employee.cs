using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BulkUpload.Models
{
    public class Employee
    {
        public int Id { get; set; }
        public string EmpName { get; set; }
        public string Position { get; set; }
        public string Location { get; set; }
        public int Age { get; set; }
        public int Salary { get; set; }
    }
}