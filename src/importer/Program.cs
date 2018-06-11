using GovUk.Education.ManageCourses.ApiClient;
using GovUk.Education.ManageCourses.Xls;
using Microsoft.Extensions.CommandLineUtils;

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

            var courses = new XlsReader().ReadCourses(folderOption.Value());
            new ManageApi().SendToManageCoursesApi(courses);
        }
   }
}
