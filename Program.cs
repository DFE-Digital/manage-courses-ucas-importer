using System;
using System.IO;
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
            foreach (var file in folder.GetFiles("*.xls"))
            {
                Console.WriteLine(" - " + file.FullName);
                using (var stream = new FileStream(file.FullName, FileMode.Open))
                {
                    var wb = new HSSFWorkbook(stream);
                    var sheet = wb.GetSheetAt(0);
                    Console.WriteLine(" -- " + sheet.SheetName);
                }
            }
        }
    }
}
