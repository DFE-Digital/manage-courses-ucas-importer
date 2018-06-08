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

            ReadFiles(folderOption.Value());
            HitApi();
        }

        private static void HitApi()
        {
            Console.WriteLine("Posting to api...");
            var client = new ManageCoursesApiClient();
            var payload = new Payload
            {
                Courses = new ObservableCollection<Course>
                {
                    new Course
                    {
                        Title = "Defence Against The Dark Arts",
                        CourseCode = "HP1",
                        NctlId = "123",
                    }
                }
            };
            client.ImportAsync(payload).Wait();
            Console.WriteLine("Done.");
        }

        private static void ReadFiles(string folder)
        {
            Console.Write("Reading xls files from: ");
            Console.WriteLine(folder);
            int totalrowcount = 0;
            foreach (var file in new DirectoryInfo(folder).GetFiles("*.xls"))
            {
                Console.WriteLine(" - " + file.FullName);
                using (var stream = new FileStream(file.FullName, FileMode.Open))
                {
                    var wb = new HSSFWorkbook(stream);
                    var sheet = wb.GetSheetAt(0);
                    Console.WriteLine(" -- " + sheet.SheetName);
                    var header = sheet.GetRow(0);
                    Console.WriteLine(" --- Cols: "
                                      + string.Join(", ", header.Cells.Select(c => c.StringCellValue)));
                    int rowcount = 0;
                    foreach (var row in sheet)
                    {
                        rowcount++;
                    }

                    Console.WriteLine(" --- rowcount: " + rowcount);
                    Console.WriteLine();
                    totalrowcount += rowcount;
                }
            }

            Console.WriteLine("total rowcount: " + totalrowcount);
        }
    }
}
