using System;
using System.Collections.ObjectModel;
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
            var urlOption = app.Option("-u|--url|--api-url <url>", "URL of the ManageCourses API", CommandOptionType.SingleValue);
            var apiKeyOption = app.Option("-k|--key|--api-key <key>", "Admin Key for the ManageCourses API", CommandOptionType.SingleValue);
            
            app.HelpOption("-?|-h|--help");
            app.Execute(args);

            if (!folderOption.HasValue()) {
                Console.WriteLine("Missing --folder option");
            }

            if (!urlOption.HasValue()) {
                Console.WriteLine("Missing --url option");
            }

            if (!apiKeyOption.HasValue()) {
                Console.WriteLine("Missing --key option");
            }

            if (!folderOption.HasValue() || !urlOption.HasValue() || !apiKeyOption.HasValue()) {
                return;
            }

            // only used to avoid importing orphaned campuses
            // i.e. we do not import institutions but need them to determine which campuses to import
            var institutions = new XlsReader().ReadInstitutions(folderOption.Value());
            
            // data to import
            var campuses = new XlsReader().ReadCampuses(folderOption.Value(), institutions);
            var courses = new XlsReader().ReadCourses(folderOption.Value(), campuses);
            var subjects = new XlsReader().ReadSubjects(folderOption.Value());
            var courseSubjects = new XlsReader().ReadCourseSubjects(folderOption.Value(), courses, subjects);
            var courseNotes = new XlsReader().ReadCourseNotes(folderOption.Value());
            var noteTexts = new XlsReader().ReadNoteText(folderOption.Value());

            var payload = new UcasPayload
            {
                Courses = new ObservableCollection<UcasCourse>(courses),
                CourseSubjects = new ObservableCollection<UcasCourseSubject>(courseSubjects),
                Subjects = new ObservableCollection<UcasSubject>(subjects),
                Campuses = new ObservableCollection<UcasCampus>(campuses),
                CourseNotes = new ObservableCollection<UcasCourseNote>(courseNotes),
                NoteTexts = new ObservableCollection<UcasNoteText>(noteTexts)
            };
            
            new ManageApi(urlOption.Value(), apiKeyOption.Value()).PostPayload(payload);
        }
    }
}
