using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel.Attributes;

namespace EmployeeReport.Models
{
    //Задачи
    public class EmpTasks
    {
        
        [ExcelColumn("ИД задачи")]
        public long ID_Task { get; set; }

        [ExcelColumn("Табельный номер")]
        public long PersonnelNumber { get; set; }
    }
}
