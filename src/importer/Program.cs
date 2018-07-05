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
            app.HelpOption("-?|-h|--help");
            app.Execute(args);

            var users = new CsvReader().ReadUsers(folderOption.Value());
            var organisations = new CsvReader().ReadOrganisations(folderOption.Value());
            var organisationInstitutions = new CsvReader().ReadOrganisationInstitutions(folderOption.Value(), organisations);
            var organisationUsers = new CsvReader().ReadOrganisationUsers(folderOption.Value(), organisations, users);
            var institutions = new XlsReader().ReadInstitutions(folderOption.Value());
            var campuses = new XlsReader().ReadCampuses(folderOption.Value(), institutions);
            var courses = new XlsReader().ReadCourses(folderOption.Value(), campuses);
            var courseSubjects = new XlsReader().ReadCourseSubjects(folderOption.Value(), courses);
            var subjects = new XlsReader().ReadSubjects(folderOption.Value());
            var courseNotes = new XlsReader().ReadCourseNotes(folderOption.Value());
            var noteTexts = new XlsReader().ReadNoteText(folderOption.Value());
            var providerMappers = new XlsReader().ReadProviderMappers(folderOption.Value());

            var payload = new Payload
            {
                Courses = new ObservableCollection<UcasCourse>(courses),
                Institutions = new ObservableCollection<UcasInstitution>(institutions),
                CourseSubjects = new ObservableCollection<UcasCourseSubject>(courseSubjects),
                Subjects = new ObservableCollection<UcasSubject>(subjects),
                Campuses = new ObservableCollection<UcasCampus>(campuses),
                CourseNotes = new ObservableCollection<UcasCourseNote>(courseNotes),
                NoteTexts = new ObservableCollection<UcasNoteText>(noteTexts),
                Users = new ObservableCollection<McUser>(users),
                Organisations = new ObservableCollection<McOrganisation>(organisations),
                OrganisationInstitutions = new ObservableCollection<McOrganisationInstitution>(organisationInstitutions),
                OrganisationUsers = new ObservableCollection<McOrganisationUser>(organisationUsers),
                Mappers = new ObservableCollection<ProviderMapper>(providerMappers)
            };
            new ManageApi().SendToManageCoursesApi(payload);
        }
    }
}
