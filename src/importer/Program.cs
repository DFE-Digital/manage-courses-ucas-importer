using System;
using System.Collections.ObjectModel;
using System.IO;
using GovUk.Education.ManageCourses.ApiClient;
using GovUk.Education.ManageCourses.Xls;
using Microsoft.Extensions.CommandLineUtils;
using Microsoft.Extensions.Configuration;

namespace GovUk.Education.ManageCourses.UcasCourseImporter
{
    class Program
    {
        static void Main(string[] args)
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .AddEnvironmentVariables()
                .AddUserSecrets<Program>()
                .Build();              

            var folder = config["folder"];            

            var ucasZipDownloader = new UcasZipDownloader();
            ucasZipDownloader.DownloadLatestToFolder(folder); 

            // only used to avoid importing orphaned campuses
            // i.e. we do not import institutions but need them to determine which campuses to import
            var institutions = new XlsReader().ReadInstitutions(folder);
            
            // data to import
            var campuses = new XlsReader().ReadCampuses(folder, institutions);
            var courses = new XlsReader().ReadCourses(folder, campuses);
            var subjects = new XlsReader().ReadSubjects(folder);
            var courseSubjects = new XlsReader().ReadCourseSubjects(folder, courses, subjects);
            var courseNotes = new XlsReader().ReadCourseNotes(folder);
            var noteTexts = new XlsReader().ReadNoteText(folder);

            var payload = new UcasPayload
            {
                Courses = new ObservableCollection<UcasCourse>(courses),
                CourseSubjects = new ObservableCollection<UcasCourseSubject>(courseSubjects),
                Subjects = new ObservableCollection<UcasSubject>(subjects),
                Campuses = new ObservableCollection<UcasCampus>(campuses),
                CourseNotes = new ObservableCollection<UcasCourseNote>(courseNotes),
                NoteTexts = new ObservableCollection<UcasNoteText>(noteTexts)
            };
            
            new ManageApi(config["url"], config["apikey"]).PostPayload(payload);
        }
    }
}
