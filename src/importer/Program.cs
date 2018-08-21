using System;
using System.Collections.ObjectModel;
using System.IO;
using GovUk.Education.ManageCourses.ApiClient;
using GovUk.Education.ManageCourses.Xls;
using Microsoft.Extensions.Configuration;
using Serilog;

namespace GovUk.Education.ManageCourses.UcasCourseImporter
{
    class Program
    {
        static void Main(string[] args)
        {
            var configuration = GetConfiguration();

            var logger = new LoggerConfiguration()
                .ReadFrom.Configuration(configuration)
                .WriteTo
                .ApplicationInsightsTraces(configuration["APPINSIGHTS_INSTRUMENTATIONKEY"])
                .CreateLogger();

            var folder = Path.Combine(Path.GetTempPath(), "ucasfiles", Guid.NewGuid().ToString());
            Directory.CreateDirectory(folder);

            var ucasZipDownloader = new UcasZipDownloader(configuration["azure_url"], configuration["azure_signature"]);
            var zipFile = ucasZipDownloader.DownloadLatestToFolder(folder).Result;

            var extractor = new UcasZipExtractor();
            var unzipFolder = Path.Combine(folder, "unzip");
            extractor.Extract(zipFile, unzipFolder);


            var xlsReader = new XlsReader(logger);

            // only used to avoid importing orphaned data
            // i.e. we do not import institutions but need them to determine which campuses to import
            var subjects = xlsReader.ReadSubjects("data");

            // data to import
            var institutions = xlsReader.ReadInstitutions(unzipFolder);
            var campuses = xlsReader.ReadCampuses(unzipFolder, institutions);
            var courses = xlsReader.ReadCourses(unzipFolder, campuses);
            var courseSubjects = xlsReader.ReadCourseSubjects(unzipFolder, courses, subjects);
            var courseNotes = xlsReader.ReadCourseNotes(unzipFolder);
            var noteTexts = xlsReader.ReadNoteText(unzipFolder);

            var payload = new UcasPayload
            {
                Institutions = new ObservableCollection<UcasInstitution>(institutions),
                Courses = new ObservableCollection<UcasCourse>(courses),
                CourseSubjects = new ObservableCollection<UcasCourseSubject>(courseSubjects),
                Campuses = new ObservableCollection<UcasCampus>(campuses),
                CourseNotes = new ObservableCollection<UcasCourseNote>(courseNotes),
                NoteTexts = new ObservableCollection<UcasNoteText>(noteTexts)
            };

            var manageApi = new ManageApi(logger, configuration["manage_api_url"], configuration["manage_api_key"]);
            manageApi.PostPayload(payload);
        }

        private static IConfiguration GetConfiguration()
        {
            return new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddJsonFile($"appsettings.{Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT") ?? "Production"}.json", optional: true)
                .AddUserSecrets<Program>()
                .AddEnvironmentVariables()
                .Build();
        }
    }
}
