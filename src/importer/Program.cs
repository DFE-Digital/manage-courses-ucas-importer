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

            var folder = Path.Combine(Path.GetTempPath(), "ucasfiles", Guid.NewGuid().ToString());
            Directory.CreateDirectory(folder);

            var ucasZipDownloader = new UcasZipDownloader(config["azure_url"], config["azure_signature"]);
            var zipFile = ucasZipDownloader.DownloadLatestToFolder(folder).Result;

            var extractor = new UcasZipExtractor();
            var unzipFolder = Path.Combine(folder, "unzip");
            extractor.Extract(zipFile, unzipFolder);


            // only used to avoid importing orphaned data
            // i.e. we do not import institutions but need them to determine which campuses to import
            var institutions = new XlsReader().ReadInstitutions(unzipFolder);
            var subjects = new XlsReader().ReadSubjects("data");

            // data to import
            var campuses = new XlsReader().ReadCampuses(unzipFolder, institutions);
            var courses = new XlsReader().ReadCourses(unzipFolder, campuses);
            var courseSubjects = new XlsReader().ReadCourseSubjects(unzipFolder, courses, subjects);
            var courseNotes = new XlsReader().ReadCourseNotes(unzipFolder);
            var noteTexts = new XlsReader().ReadNoteText(unzipFolder);

            var payload = new UcasPayload
            {
                Courses = new ObservableCollection<UcasCourse>(courses),
                CourseSubjects = new ObservableCollection<UcasCourseSubject>(courseSubjects),
                Campuses = new ObservableCollection<UcasCampus>(campuses),
                CourseNotes = new ObservableCollection<UcasCourseNote>(courseNotes),
                NoteTexts = new ObservableCollection<UcasNoteText>(noteTexts)
            };

            new ManageApi(config["manage_api_url"], config["manage_api_key"]).PostPayload(payload);
        }
    }
}
