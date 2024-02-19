using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel.Attributes;

namespace EmployeeReport.Models
{
    //Отделы
   public class Departments
    {
        [ExcelColumn("ИД отдела")]
        public int ID_Department  {get; set; }//ID отдела
        [ExcelColumn("Наименование отдела")]
        public string DepartmentName { get; set; }//Наименование отдела
    }
}
