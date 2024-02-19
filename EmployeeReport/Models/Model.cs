using System;
using System.Collections.Generic;
using System.IO;
using LinqToExcel;
using System.Linq;
using LinqToExcel.Query;

namespace EmployeeReport.Models
{
    public class Model
    {
        static Model _instance;
        readonly List<Employee> _employee = new List<Employee>();
        readonly List<Departments> _departments = new List<Departments>();
        readonly List<EmpTasks> _tasks = new List<EmpTasks>();
        private ExcelQueryFactory _excel;
        Model()
        {
            _excel = new ExcelQueryFactory(@"Data.xlsb")
            {
                DatabaseEngine = LinqToExcel.Domain.DatabaseEngine.Jet,
                TrimSpaces = LinqToExcel.Query.TrimSpacesType.End,
                UsePersistentConnection = true,
                ReadOnly = true

            };
            var Employes = from p in _excel.Worksheet<Employee>("Сотрудники")
                           select p;

            foreach (Employee employee in Employes)
            {
                _employee.Add(new Employee { PersonnelNumber = employee.PersonnelNumber, SurName = employee.SurName, Name = employee.Name, Patronymic = employee.Patronymic, DateBirth = employee.DateBirth, Department = employee.Department });
            }

            var Department = from z in _excel.Worksheet<Departments>("Отделы")
                             select z;

            foreach (Departments department in Department)
            {
                _departments.Add(new Departments { ID_Department = department.ID_Department, DepartmentName = department.DepartmentName });
            }

            var Task = from y in _excel.Worksheet<EmpTasks>("Задачи")
                             select y;

            foreach (EmpTasks task in Task)
            {
                _tasks.Add(new EmpTasks { ID_Task = task.ID_Task, PersonnelNumber = task.PersonnelNumber});
            }
        }
        public static Model getInstance()
        {
            if (_instance == null)
                _instance = new Model();
            return _instance;
        }

        public List<Employee> GetEmployee()
        {
            return _employee.ToList();
        }

        public List<Departments> GetDepartments()
        {
            return _departments.ToList();
        }

        public List<EmpTasks> GetTasks()
        {
            return _tasks.ToList();
        }
    }
}
