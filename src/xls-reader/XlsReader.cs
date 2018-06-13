using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using GovUk.Education.ManageCourses.ApiClient;
using NPOI.HSSF.UserModel;

namespace GovUk.Education.ManageCourses.Xls
{
    public class XlsReader
    {
        public  List<UcasCourse> ReadCourses(string folder)
        {
            var file = new FileInfo(Path.Combine(folder, "GTTR_CRSE.xls"));
            Console.WriteLine("Reading course xls file from: " + file.FullName);

            var courses = new List<UcasCourse>();
            using (var stream = new FileStream(file.FullName, FileMode.Open))
            {
                var wb = new HSSFWorkbook(stream);
                var sheet = wb.GetSheetAt(0);
                var header = sheet.GetRow(0);
                
                var columnMap = header.Cells.ToDictionary(c => c.StringCellValue, c => c.ColumnIndex);
                for (int dataRowIndex = 1; dataRowIndex <= sheet.LastRowNum; dataRowIndex++)
                {
                    if (sheet.GetRow(dataRowIndex) == null)
                    {
                        continue;
                    }

                    var row = sheet.GetRow(dataRowIndex);
                    courses.Add(new UcasCourse
                        {
                            Title = row.GetCell(columnMap["CRSE_TITLE"]).StringCellValue,
                            CourseCode = row.GetCell(columnMap["CRSE_CODE"]).StringCellValue,
                            InstCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue,
                            StudyMode = row.GetCell(columnMap["STUDYMODE"]).StringCellValue,
                            AgeGroup = row.GetCell(columnMap["AGE"]).StringCellValue,
                            CampusCode = row.GetCell(columnMap["CAMPUS_CODE"]).StringCellValue,
                            ProfPostFlag = row.GetCell(columnMap["PROFPOST_FLAG"]).StringCellValue,
                            ProgramType = row.GetCell(columnMap["PROGRAM_TYPE"]).StringCellValue,
                            AcreditedProvider = row.GetCell(columnMap["ACCREDITING_PROVIDER"]).StringCellValue,
                            OpenDate = row.GetCell(columnMap["CRSE_OPEN_DATE"]).StringCellValue,
                        }
                    );
                }
            }
            Console.Out.WriteLine(courses.Count + " courses loaded from xls");
            return courses;
        }
        public List<UcasInstitution> ReadInstitutions(string folder)
        {
            var file = new FileInfo(Path.Combine(folder, "GTTR_INST.xls"));
            Console.WriteLine("Reading institution xls file from: " + file.FullName);

            var institutions = new List<UcasInstitution>();
            using (var stream = new FileStream(file.FullName, FileMode.Open))
            {
                var wb = new HSSFWorkbook(stream);
                var sheet = wb.GetSheetAt(0);
                var header = sheet.GetRow(0);

                var columnMap = header.Cells.ToDictionary(c => c.StringCellValue, c => c.ColumnIndex);
                for (int dataRowIndex = 1; dataRowIndex <= sheet.LastRowNum; dataRowIndex++)
                {
                    if (sheet.GetRow(dataRowIndex) == null)
                    {
                        continue;
                    }

                    var row = sheet.GetRow(dataRowIndex);
                    institutions.Add(new UcasInstitution
                        {
                            InstCode = row.GetCell(columnMap["INST_NAME"]).StringCellValue,
                            FullName = row.GetCell(columnMap["INST_FULL"]).StringCellValue,
                            InstType = row.GetCell(columnMap["INST_TYPE"]).StringCellValue,
                            Address1 = row.GetCell(columnMap["ADDR_1"]).StringCellValue,
                            Address2 = row.GetCell(columnMap["ADDR_2"]).StringCellValue,
                            Address3 = row.GetCell(columnMap["ADDR_3"]).StringCellValue,
                            Address4 = row.GetCell(columnMap["ADDR_4"]).StringCellValue,
                            PostCode = row.GetCell(columnMap["POSTCODE"]).StringCellValue,
                            ContactName = row.GetCell(columnMap["CONTACT_NAME"]).StringCellValue,
                            Url = row.GetCell(columnMap["URL"]).StringCellValue,
                            Scitt = row.GetCell(columnMap["SCITT"]).StringCellValue,
                            AcreditedProvider = row.GetCell(columnMap["ACCREDITING_PROVIDER"]).StringCellValue,
                        }
                    );
                }
            }
            Console.Out.WriteLine(institutions.Count + " intitutions loaded from xls");
            return institutions;
        }
        public List<UcasSubject> ReadSubjects(string folder)
        {
            var file = new FileInfo(Path.Combine(folder, "GTTR_CRSE_SUBJECT.xls"));
            Console.WriteLine("Reading subject xls file from: " + file.FullName);

            var subjects = new List<UcasSubject>();
            using (var stream = new FileStream(file.FullName, FileMode.Open))
            {
                var wb = new HSSFWorkbook(stream);
                var sheet = wb.GetSheetAt(0);
                var header = sheet.GetRow(0);

                var columnMap = header.Cells.ToDictionary(c => c.StringCellValue, c => c.ColumnIndex);
                for (int dataRowIndex = 1; dataRowIndex <= sheet.LastRowNum; dataRowIndex++)
                {
                    if (sheet.GetRow(dataRowIndex) == null)
                    {
                        continue;
                    }

                    var row = sheet.GetRow(dataRowIndex);
                    subjects.Add(new UcasSubject
                    {
                        InstCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue,
                        CourseCode = row.GetCell(columnMap["CRSE_CODE"]).StringCellValue,
                        SubjectCode = row.GetCell(columnMap["SUBJECT_CODE"]).StringCellValue,
                    }
                    );
                }
            }
            Console.Out.WriteLine(subjects.Count + " subjects loaded from xls");
            return subjects;
        }
        public List<UcasCampus> ReadCampuses(string folder)
        {
            var file = new FileInfo(Path.Combine(folder, "GTTR_CAMPUS.xls"));
            Console.WriteLine("Reading campus xls file from: " + file.FullName);

            var campuses = new List<UcasCampus>();
            using (var stream = new FileStream(file.FullName, FileMode.Open))
            {
                var wb = new HSSFWorkbook(stream);
                var sheet = wb.GetSheetAt(0);
                var header = sheet.GetRow(0);

                var columnMap = header.Cells.ToDictionary(c => c.StringCellValue, c => c.ColumnIndex);
                for (int dataRowIndex = 1; dataRowIndex <= sheet.LastRowNum; dataRowIndex++)
                {
                    if (sheet.GetRow(dataRowIndex) == null)
                    {
                        continue;
                    }

                    var row = sheet.GetRow(dataRowIndex);
                    campuses.Add(new UcasCampus
                        {
                            InstCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue,
                            CampusCode = row.GetCell(columnMap["CAMPUS_CODE"]).StringCellValue,
                            CampusName = row.GetCell(columnMap["CAMPUS_NAME"]).StringCellValue,
                            Address1 = row.GetCell(columnMap["ADDR_1"]).StringCellValue,
                            Address2 = row.GetCell(columnMap["ADDR_2"]).StringCellValue,
                            Address3 = row.GetCell(columnMap["ADDR_3"]).StringCellValue,
                            Address4 = row.GetCell(columnMap["ADDR_4"]).StringCellValue,
                            PostCode = row.GetCell(columnMap["POSTCODE"]).StringCellValue,
                            TelNo = row.GetCell(columnMap["TEL_NO"]).StringCellValue,
                            RegionCode = row.GetCell(columnMap["REGION_CODE"]).StringCellValue,
                    }
                    );
                }
            }
            Console.Out.WriteLine(campuses.Count + " campuses loaded from xls");
            return campuses;
        }
        public List<UcasCourseNote> ReadCourseNotes(string folder)
        {
            var file = new FileInfo(Path.Combine(folder, "GTTR_CRSENOTE.xls"));
            Console.WriteLine("Reading course note xls file from: " + file.FullName);

            var courseNotes = new List<UcasCourseNote>();
            using (var stream = new FileStream(file.FullName, FileMode.Open))
            {
                var wb = new HSSFWorkbook(stream);
                var sheet = wb.GetSheetAt(0);
                var header = sheet.GetRow(0);

                var columnMap = header.Cells.ToDictionary(c => c.StringCellValue, c => c.ColumnIndex);
                for (int dataRowIndex = 1; dataRowIndex <= sheet.LastRowNum; dataRowIndex++)
                {
                    if (sheet.GetRow(dataRowIndex) == null)
                    {
                        continue;
                    }

                    var row = sheet.GetRow(dataRowIndex);
                    courseNotes.Add(new UcasCourseNote
                        {
                            InstCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue,
                            CourseCode = row.GetCell(columnMap["CRSE_CODE"]).StringCellValue,
                            NoteNo = row.GetCell(columnMap["NOTE_NO"]).StringCellValue,
                            NoteType = row.GetCell(columnMap["NOTE_TYPE"]).StringCellValue,
                    }
                    );
                }
            }
            Console.Out.WriteLine(courseNotes.Count + " course notes loaded from xls");
            return courseNotes;
        }
        public List<UcasNoteText> ReadNoteText(string folder)
        {
            var file = new FileInfo(Path.Combine(folder, "GTTR_NOTETEXT.xls"));
            Console.WriteLine("Reading note text xls file from: " + file.FullName);

            var noteTexts = new List<UcasNoteText>();
            using (var stream = new FileStream(file.FullName, FileMode.Open))
            {
                var wb = new HSSFWorkbook(stream);
                var sheet = wb.GetSheetAt(0);
                var header = sheet.GetRow(0);

                var columnMap = header.Cells.ToDictionary(c => c.StringCellValue, c => c.ColumnIndex);
                for (int dataRowIndex = 1; dataRowIndex <= sheet.LastRowNum; dataRowIndex++)
                {
                    if (sheet.GetRow(dataRowIndex) == null)
                    {
                        continue;
                    }

                    var row = sheet.GetRow(dataRowIndex);
                    noteTexts.Add(new UcasNoteText
                        {
                            InstCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue,
                            NoteNo = row.GetCell(columnMap["NOTE_NO"]).StringCellValue,
                            NoteType = row.GetCell(columnMap["NOTE_TYPE"]).StringCellValue,
                            LineText = row.GetCell(columnMap["LINE_TEXT"]).StringCellValue,
                        }
                    );
                }
            }
            Console.Out.WriteLine(noteTexts.Count + " note texts loaded from xls");
            return noteTexts;
        }
    }
}
