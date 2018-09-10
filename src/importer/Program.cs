using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CsvHelper;
using GovUk.Education.ManageCourses.ApiClient;
using GovUk.Education.ManageCourses.Csv.Domain;
using GovUk.Education.ManageCourses.Xls;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.DependencyCollector;
using Microsoft.ApplicationInsights.Extensibility;
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

            logger.Information("UcasCourseImporter started.");

            var folder = Path.Combine(Path.GetTempPath(), "ucasfiles", Guid.NewGuid().ToString());
            logger.Information($"Using folder {folder}");
            Directory.CreateDirectory(folder);

            var downloadAndExtractor = new DownloaderAndExtractor(logger, folder, configuration["azure_url"], configuration["azure_signature"]);

            var unzipFolder = downloadAndExtractor.DownloadAndExtractLatest("NetupdateExtract");
            var unzipFolderProfiles = downloadAndExtractor.DownloadAndExtractLatest("EntryProfilesExtract_test");

            var xlsReader = new XlsReader(logger);
            
            // only used to avoid importing orphaned data
            // i.e. we do not import institutions but need them to determine which campuses to import
            var subjects = xlsReader.ReadSubjects("data");

            // entry profile data - used to correct institution data
            var institutionProfiles = ReadInstitutionProfiles(unzipFolderProfiles);

            // data to import
            var institutions = xlsReader.ReadInstitutions(unzipFolder);
            UpdateContactDetails(institutions, institutionProfiles);

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

            logger.Information("UcasCourseImporter finished.");
        }

        private static Dictionary<string, UcasInstitutionProfile> ReadInstitutionProfiles(string unzipFolderProfiles)
        {
            var institutionProfiles = new Dictionary<string, UcasInstitutionProfile>();
            var institutionProfilesCsv = new CsvReader(File.OpenText(Path.Combine(unzipFolderProfiles, "gttr_inst.csv")));
            institutionProfilesCsv.Read();
            institutionProfilesCsv.ReadHeader();
            while (institutionProfilesCsv.Read())
            {
                var rec = institutionProfilesCsv.GetRecord<UcasInstitutionProfile>();
                institutionProfiles[rec.inst_code] = rec;
            }

            return institutionProfiles;
        }

        private static void UpdateContactDetails(List<UcasInstitution> institutions, IDictionary<string, UcasInstitutionProfile> institutionProfiles)
        {
            foreach(var inst in institutions)
            {
                if (institutionProfiles.TryGetValue(inst.InstCode, out UcasInstitutionProfile profile))
                {
                    inst.Addr1 = profile.inst_address1.Trim();
                    inst.Addr2 = profile.inst_address2.Trim();
                    inst.Addr3 = profile.inst_address3.Trim();
                    inst.Addr4 = profile.inst_address4.Trim();
                    inst.Postcode = profile.inst_post_code.Trim();
                    inst.ContactName = profile.inst_person.Trim();
                    inst.Email = profile.email.Trim();
                    inst.Telephone = profile.inst_tel.Trim();
                    inst.Url = profile.web_addr.Trim();
                }
            }
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
