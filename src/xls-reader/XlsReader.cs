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
        public List<UcasCourse> ReadCourses(string folder)
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
                        InstCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue.Trim(),
                        CrseCode = row.GetCell(columnMap["CRSE_CODE"]).StringCellValue.Trim(),
                        CrseTitle = row.GetCell(columnMap["CRSE_TITLE"]).StringCellValue.Trim(),
                        Studymode = row.GetCell(columnMap["STUDYMODE"]).StringCellValue.Trim(),
                        Age = row.GetCell(columnMap["AGE"]).StringCellValue.Trim(),
                        CampusCode = row.GetCell(columnMap["CAMPUS_CODE"]).StringCellValue.Trim(),
                        ProfpostFlag = row.GetCell(columnMap["PROFPOST_FLAG"]).StringCellValue.Trim(),
                        ProgramType = row.GetCell(columnMap["PROGRAM_TYPE"]).StringCellValue.Trim(),
                        AccreditingProvider = row.GetCell(columnMap["ACCREDITING_PROVIDER"]).StringCellValue.Trim(),
                        CrseOpenDate = row.GetCell(columnMap["CRSE_OPEN_DATE"]).StringCellValue.Trim(),
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
                        InstCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue.Trim(),
                        InstName = row.GetCell(columnMap["INST_NAME"]).StringCellValue.Trim(),
                        InstBig = row.GetCell(columnMap["INST_BIG"]).StringCellValue.Trim(),
                        InstFull = row.GetCell(columnMap["INST_FULL"]).StringCellValue.Trim(),
                        InstType = row.GetCell(columnMap["INST_TYPE"]).StringCellValue.Trim(),
                        Addr1 = row.GetCell(columnMap["ADDR_1"]).StringCellValue.Trim(),
                        Addr2 = row.GetCell(columnMap["ADDR_2"]).StringCellValue.Trim(),
                        Addr3 = row.GetCell(columnMap["ADDR_3"]).StringCellValue.Trim(),
                        Addr4 = row.GetCell(columnMap["ADDR_4"]).StringCellValue.Trim(),
                        Postcode = row.GetCell(columnMap["POSTCODE"]).StringCellValue.Trim(),
                        ContactName = row.GetCell(columnMap["CONTACT_NAME"]).StringCellValue.Trim(),
                        YearCode = row.GetCell(columnMap["YEAR_CODE"]).StringCellValue.Trim(),
                        Url = row.GetCell(columnMap["URL"]).StringCellValue.Trim(),
                        Scitt = row.GetCell(columnMap["SCITT"]).StringCellValue.Trim(),
                        AccreditingProvider = row.GetCell(columnMap["ACCREDITING_PROVIDER"]).StringCellValue.Trim(),
                        SchemeMember = row.GetCell(columnMap["SCHEME_MEMBER"]).StringCellValue.Trim()
                    }
                    );
                }
            }
            Console.Out.WriteLine(institutions.Count + " intitutions loaded from xls");
            return institutions;
        }

        public List<UcasCourseSubject> ReadCourseSubjects(string folder, IList<UcasCourse> courses)
        {
            var file = new FileInfo(Path.Combine(folder, "GTTR_CRSE_SUBJECT.xls"));
            Console.WriteLine("Reading course subject xls file from: " + file.FullName);

            var courseSubjects = new List<UcasCourseSubject>();
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
                    var ucasCourseSubject = new UcasCourseSubject
                    {
                        InstCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue.Trim(),
                        CrseCode = row.GetCell(columnMap["CRSE_CODE"]).StringCellValue.Trim(),
                        SubjectCode = row.GetCell(columnMap["SUBJECT_CODE"]).StringCellValue.Trim(),
                        YearCode = row.GetCell(columnMap["YEAR_CODE"]).StringCellValue.Trim()
                    };
                    if (!courses.Any(c => c.InstCode == ucasCourseSubject.InstCode && c.CrseCode == ucasCourseSubject.CrseCode))
                    {
                        Console.Out.WriteLine($"UcasCourseSubject skipped - invalid inst_code/crse_code combination: {ucasCourseSubject.InstCode}, {ucasCourseSubject.CrseCode}");
                        continue;
                    }
                    courseSubjects.Add(ucasCourseSubject
                    );
                }
            }
            Console.Out.WriteLine(courseSubjects.Count + " course subjects loaded from xls");
            return courseSubjects;
        }
        public List<UcasSubject> ReadSubjects(string folder)
        {
            var file = new FileInfo(Path.Combine(folder, "GTTR_SUBJECT.xls"));
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
                        SubjectCode = row.GetCell(columnMap["SUBJECT_CODE"]).StringCellValue.Trim(),
                        SubjectDescription = row.GetCell(columnMap["SUBJECT_DESCRIPTION"]).StringCellValue.Trim(),
                        TitleMatch = row.GetCell(columnMap["TITLE_MATCH"]).StringCellValue.Trim(),
                    }
                    );
                }
            }
            Console.Out.WriteLine(subjects.Count + " subjects loaded from xls");
            return subjects;
        }
        public List<UcasCampus> ReadCampuses(string folder, IList<UcasInstitution> institutions)
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
                    var ucasCampus = new UcasCampus
                    {
                        InstCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue.Trim(),
                        CampusCode = row.GetCell(columnMap["CAMPUS_CODE"]).StringCellValue.Trim(),
                        CampusName = row.GetCell(columnMap["CAMPUS_NAME"]).StringCellValue.Trim(),
                        Addr1 = row.GetCell(columnMap["ADDR_1"]).StringCellValue.Trim(),
                        Addr2 = row.GetCell(columnMap["ADDR_2"]).StringCellValue.Trim(),
                        Addr3 = row.GetCell(columnMap["ADDR_3"]).StringCellValue.Trim(),
                        Addr4 = row.GetCell(columnMap["ADDR_4"]).StringCellValue.Trim(),
                        Postcode = row.GetCell(columnMap["POSTCODE"]).StringCellValue.Trim(),
                        TelNo = row.GetCell(columnMap["TEL_NO"]).StringCellValue.Trim(),
                        Email = row.GetCell(columnMap["EMAIL"]).StringCellValue.Trim(),
                        RegionCode = row.GetCell(columnMap["REGION_CODE"]).StringCellValue.Trim(),
                    };
                    if (!institutions.Any(i => i.InstCode == ucasCampus.InstCode))
                    {
                        Console.Out.WriteLine($"Campus '{ucasCampus.CampusCode}' skipped - invalid inst_code {ucasCampus.InstCode}");
                        continue;
                    }
                    campuses.Add(ucasCampus);
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
                        InstCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue.Trim(),
                        CrseCode = row.GetCell(columnMap["CRSE_CODE"]).StringCellValue.Trim(),
                        NoteNo = row.GetCell(columnMap["NOTE_NO"]).StringCellValue.Trim(),
                        NoteType = row.GetCell(columnMap["NOTE_TYPE"]).StringCellValue.Trim(),
                        YearCode = row.GetCell(columnMap["YEAR_CODE"]).StringCellValue.Trim()
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
                        InstCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue.Trim(),
                        NoteNo = row.GetCell(columnMap["NOTE_NO"]).StringCellValue.Trim(),
                        NoteType = row.GetCell(columnMap["NOTE_TYPE"]).StringCellValue.Trim(),
                        LineText = row.GetCell(columnMap["LINE_TEXT"]).StringCellValue.Trim(),
                        YearCode = row.GetCell(columnMap["YEAR_CODE"]).StringCellValue.Trim()
                    }
                    );
                }
            }
            Console.Out.WriteLine(noteTexts.Count + " note texts loaded from xls");
            return noteTexts;
        }
        public List<ProviderMapper> ReadProviderMappers(string folder)
        {
            var file = new FileInfo(Path.Combine(folder, "ProviderMapper.xls"));
            Console.WriteLine("Reading provider mapper xls file from: " + file.FullName);

            var providerMappers = new List<ProviderMapper>();
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
                    providerMappers.Add(new ProviderMapper
                    {
                        OrgId = row.GetCell(columnMap["ORG_ID"]).NumericCellValue.ToString(),
                        UcasCode = row.GetCell(columnMap["UCAS_CODE"]).StringCellValue.Trim(),
                        Urn = (int)row.GetCell(columnMap["URN"]).NumericCellValue,
                        Type = row.GetCell(columnMap["Type"]).StringCellValue.Trim(),
                        InstitutionName = row.GetCell(columnMap["Institution_Name"]).StringCellValue.Trim(),
                    }
                    );
                }
            }
            Console.Out.WriteLine(providerMappers.Count + " note texts loaded from xls");
            return providerMappers;
        }
    }
}
