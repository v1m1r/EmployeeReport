using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EmployeeReport.Models;

namespace EmployeeReport.ViewModel
{
    public class EmployeeReportViewModel
    {
        public Employee Employee { get; set; }
        public Departments Departments { get; set; }

        public EmpTasks EmpTasks { get; set; }
    }
}
