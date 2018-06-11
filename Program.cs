using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using Microsoft.Extensions.CommandLineUtils;
using NPOI.HSSF.UserModel;

namespace GovUk.Education.ManageCourses.UcasCourseImporter
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new CommandLineApplication();
            var folderOption = app.Option("-$|-f|--folder <folder>", "Folder to read UCAS .xls files from ", CommandOptionType.SingleValue);
            app.HelpOption("-?|-h|--help");
            app.Execute(args);

            var courses = ReadCourses(folderOption.Value());
            SendToManageCoursesApi(courses);
        }

        private static void SendToManageCoursesApi(List<Course> courses)
        {
            Console.WriteLine("Posting to api...");
            var client = new ManageCoursesApiClient();
            var payload = new Payload { Courses = new ObservableCollection<Course>(courses) };
            client.ImportAsync(payload).Wait();
            Console.WriteLine("Done.");
        }

        private static List<Course> ReadCourses(string folder)
        {
            Console.Write("Reading course xls file from: ");
            Console.WriteLine(folder);
            var courses = new List<Course>();
            var file = new FileInfo(Path.Combine(folder, "GTTR_CRSE.xls"));
            using (var stream = new FileStream(file.FullName, FileMode.Open))
            {
                var wb = new HSSFWorkbook(stream);
                var sheet = wb.GetSheetAt(0);
                var header = sheet.GetRow(0);
                var columnMap = header.Cells.ToDictionary(c => c.StringCellValue, c => c.ColumnIndex);
                for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    if (sheet.GetRow(rowIndex) == null)
                    {
                        continue;
                    }

                    var row = sheet.GetRow(rowIndex);
                    courses.Add(new Course
                    {
                        UcasInstitutionCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue,
                        UcasCourseCode = row.GetCell(columnMap["CRSE_CODE"]).StringCellValue,
                        Title = row.GetCell(columnMap["CRSE_TITLE"]).StringCellValue,
                    });
                }
            }
            Console.Out.WriteLine(courses.Count + " courses loaded from xls");
            return courses;
        }
    }
}
