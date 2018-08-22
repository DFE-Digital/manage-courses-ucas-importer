using System;
using System.Collections.ObjectModel;
using System.IO;
using GovUk.Education.ManageCourses.ApiClient;
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

            TelemetryConfiguration.Active.InstrumentationKey = configuration["APPINSIGHTS_INSTRUMENTATIONKEY"];
            var telemetryClient = new TelemetryClient(TelemetryConfiguration.Active);

            // boilerplate code from https://docs.microsoft.com/en-us/azure/application-insights/application-insights-console
            var module = new DependencyTrackingTelemetryModule();
            module.ExcludeComponentCorrelationHttpHeadersOnDomains.Add("core.windows.net"); // prevent Correlation Id to be sent to certain endpoints. You may add other domains as needed.
            // enable known dependency tracking, note that in future versions, we will extend this list. 
            // please check default settings in https://github.com/Microsoft/ApplicationInsights-dotnet-server/blob/develop/Src/DependencyCollector/NuGet/ApplicationInsights.config.install.xdt#L20
            module.IncludeDiagnosticSourceActivities.Add("Microsoft.Azure.ServiceBus");
            module.IncludeDiagnosticSourceActivities.Add("Microsoft.Azure.EventHubs");
            module.Initialize(TelemetryConfiguration.Active); // initialize the module
            TelemetryConfiguration.Active.TelemetryInitializers.Add(new OperationCorrelationTelemetryInitializer()); // stamps telemetry with correlation identifiers
            TelemetryConfiguration.Active.TelemetryInitializers.Add(new HttpDependenciesParsingTelemetryInitializer()); // ensures proper DependencyTelemetry.Type is set for Azure RESTful API calls

            telemetryClient.TrackEvent("TestEvent");

            var logger = new LoggerConfiguration()
                .ReadFrom.Configuration(configuration)
                .WriteTo
                .ApplicationInsightsTraces(telemetryClient)
                .CreateLogger();

            logger.Information("UcasCourseImporter started.");

            var folder = Path.Combine(Path.GetTempPath(), "ucasfiles", Guid.NewGuid().ToString());
            logger.Information($"Using folder {folder}");
            Directory.CreateDirectory(folder);

            var ucasZipDownloader = new UcasZipDownloader(logger, configuration["azure_url"], configuration["azure_signature"]);
            var zipFile = ucasZipDownloader.DownloadLatestToFolder(folder).Result;

            var unzipFolder = Path.Combine(folder, "unzip");
            logger.Information($"Unzipping {zipFile} to {unzipFolder}");
            var extractor = new UcasZipExtractor();
            extractor.Extract(zipFile, unzipFolder);

            telemetryClient.Flush();

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
            telemetryClient.Flush();
            manageApi.PostPayload(payload);

            logger.Information("UcasCourseImporter finished.");
            telemetryClient.Flush();
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
