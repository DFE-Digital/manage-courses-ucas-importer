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
        public  List<Course> ReadCourses(string folder)
        {
            var file = new FileInfo(Path.Combine(folder, "GTTR_CRSE.xls"));
            Console.WriteLine("Reading course xls file from: " + file.FullName);

            var courses = new List<Course>();
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
                    courses.Add(new Course
                    {
                        //UcasInstitutionCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue,
                        UcasCode = row.GetCell(columnMap["CRSE_CODE"]).StringCellValue,
                        Title = row.GetCell(columnMap["CRSE_TITLE"]).StringCellValue,
                        //build type field so as to contain all required fields                        
                        Type = row.GetCell(columnMap["INST_CODE"]).StringCellValue + "," +
                               row.GetCell(columnMap["STUDYMODE"]).StringCellValue + "," +
                               row.GetCell(columnMap["AGE"]).StringCellValue + "," +
                               row.GetCell(columnMap["PROFPOST_FLAG"]).StringCellValue + "," +
                               row.GetCell(columnMap["COPY_FORM_REQD"]).StringCellValue + "," +
                               row.GetCell(columnMap["PROGRAM_TYPE"]).StringCellValue + "," +
                               row.GetCell(columnMap["HAS_BEEN_PUBLISHED"]).StringCellValue + "," +
                               row.GetCell(columnMap["ACCREDITING_PROVIDER"]).StringCellValue + "," +
                               row.GetCell(columnMap["CRSE_OPEN_DATE"]).StringCellValue
                    });
                }
            }
            Console.Out.WriteLine(courses.Count + " courses loaded from xls");
            return courses;
        }
    }
}
