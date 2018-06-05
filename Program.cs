using System;
using System.IO;
using Microsoft.Extensions.CommandLineUtils;

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
            }
        }
    }
}
