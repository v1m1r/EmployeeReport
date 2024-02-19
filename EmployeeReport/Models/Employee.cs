using LinqToExcel.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmployeeReport.Models
{
    //Сотрудники
    public class Employee
    {
        [ExcelColumn("Табельный номер")]
        public long PersonnelNumber { get; set; }//Табельный номер

        [ExcelColumn("Фамилия")]
        public string SurName { get; set; } //Фамилия

        [ExcelColumn("Имя")]
        public string Name { get; set; } //Имя

        [ExcelColumn("Отчество")]
        public string Patronymic { get; set; } //Отчество

        [ExcelColumn("Дата рождения")]
        public DateTime DateBirth { get; set; }//Дата рождения

        [ExcelColumn("Отдел")]
        public int Department { get; set; }//Отдел
    }
}
