using System;
using System.IO;
using System.Linq;
using Microsoft.Extensions.CommandLineUtils;
using NPOI.HSSF.UserModel;

namespace UcasCourseImporter
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new CommandLineApplication();
            var folderOption = app.Option("-$|-f|--folder <folder>", "Folder to read UCAS .xls files from ", CommandOptionType.SingleValue);
            app.HelpOption("-?|-h|--help");
            app.Execute(args);
            Console.Write("Reading xls files from: ");
            Console.WriteLine(folderOption.Value());

            var folder = new DirectoryInfo(folderOption.Value());
            int totalrowcount = 0;
            foreach (var file in folder.GetFiles("*.xls"))
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
