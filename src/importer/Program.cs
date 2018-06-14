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

            var courses = new XlsReader().ReadCourses(folderOption.Value());
            var institutions = new XlsReader().ReadInstitutions(folderOption.Value());
            var courseSubjects = new XlsReader().ReadCourseSubjects(folderOption.Value());
            var subjects = new XlsReader().ReadSubjects(folderOption.Value());
            var campuses = new XlsReader().ReadCampuses(folderOption.Value());
            var courseNotes = new XlsReader().ReadCourseNotes(folderOption.Value());
            var noteTexts = new XlsReader().ReadNoteText(folderOption.Value());

            var payload = new Payload
            {
                Courses = new ObservableCollection<UcasCourse>(courses),
                Institutions = new ObservableCollection<UcasInstitution>(institutions),
                CourseSubjects = new ObservableCollection<UcasCourseSubject>(courseSubjects),
                Subjects = new ObservableCollection<UcasSubject>(subjects),
                Campuses = new ObservableCollection<UcasCampus>(campuses),
                CourseNotes = new ObservableCollection<UcasCourseNote>(courseNotes),
                NoteTexts = new ObservableCollection<UcasNoteText>(noteTexts)
            };
            new ManageApi().SendToManageCoursesApi(payload);
        }
   }
}
