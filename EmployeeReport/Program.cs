using System;
using System.Data;
using System.Collections.ObjectModel;
using EmployeeReport.ViewModel;
using System.Linq;
using EmployeeReport.Models;
using Microsoft.Office.Interop.Word;
namespace EmployeeReport
{
    class Program
    {
        static Model _model;
        static ObservableCollection<EmployeeReportViewModel> EmployeeReport2 { get; set; }
        static void Main(string[] args)
        {
            try
            {
                _model = Model.getInstance();
                var _employee = _model.GetEmployee();
                var _departments = _model.GetDepartments();
                var _tasks = _model.GetTasks();

                var obj = from e in _employee
                          join d in _departments on e.Department equals d.ID_Department
                          join t in _tasks on e.PersonnelNumber equals t.PersonnelNumber
                          into g
                          select new
                          {
                              e.Name,
                              e.SurName,
                              e.Patronymic,
                              e.DateBirth,
                              d.DepartmentName,
                              t = g.Count()
                          };
                var GroupDepartment = obj.GroupBy(x => x.DepartmentName);
               
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document doc = app.Documents.Add();

                try
                {
                    Microsoft.Office.Interop.Word.Table tbl = app.ActiveDocument.Tables.Add(doc.Range(), obj.Count(), 3);
                    tbl.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                    tbl.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                    tbl.Rows[1].Cells[1].Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    tbl.Rows[1].Cells[1].Shading.BackgroundPatternColor = WdColor.wdColorGray65;
                    tbl.Rows[1].Cells[1].Range.Font.Color = WdColor.wdColorWhite;
                    tbl.Rows[1].Cells[1].Range.Text = "ФИО";
                    tbl.Rows[1].Cells[2].Shading.BackgroundPatternColor = WdColor.wdColorGray65;
                    tbl.Rows[1].Cells[2].Range.Font.Color = WdColor.wdColorWhite;
                    tbl.Rows[1].Cells[2].Range.Text = "Количество задач";
                    tbl.Rows[1].Cells[3].Shading.BackgroundPatternColor = WdColor.wdColorGray65;
                    tbl.Rows[1].Cells[3].Range.Font.Color = WdColor.wdColorWhite;
                    tbl.Rows[1].Cells[3].Range.Text = "Отдел";

                    int i = 2;
                    foreach (var grouping in GroupDepartment)
                    {
                        Console.WriteLine(grouping.Key);
                        foreach (var rec in grouping)
                        {
                            
                            tbl.Rows[i].Cells[1].Range.Text = string.Format("{0} {1}.{2}.",rec.SurName,rec.Name.Substring(0,1),rec.Patronymic.Substring(0,1));
                            tbl.Rows[i].Cells[2].Range.Text = rec.t.ToString();
                            tbl.Rows[i].Cells[3].Range.Text = grouping.Key;
                            i++;
                        }
                    }
                    tbl.Rows.Add();
                }

                catch (Exception ex) { Console.WriteLine(string.Format("При создании отчета в word возникла ошибка {0}", ex.Message)); }

                app.Visible = true;
                Console.ReadLine();
            }
            catch (Exception ex) { Console.WriteLine(string.Format("При работе с Excel возникла ошибка {0}", ex.Message)); }



        }
    }
}
