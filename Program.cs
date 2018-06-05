using System;
using Microsoft.Extensions.CommandLineUtils;

namespace UcasCourseImporter
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new CommandLineApplication();
            var folder = app.Option("-$|-f|--folder <folder>", "Folder to read UCAS .xls files from ", CommandOptionType.SingleValue);
            app.HelpOption("-?|-h|--help");
            app.Execute(args);
            Console.Write("Reading xls files from: ");
            Console.WriteLine(folder.Value());
        }
    }
}
