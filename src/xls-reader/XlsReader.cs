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
                        UcasInstitutionCode = row.GetCell(columnMap["INST_CODE"]).StringCellValue,
                        UcasCourseCode = row.GetCell(columnMap["CRSE_CODE"]).StringCellValue,
                        Title = row.GetCell(columnMap["CRSE_TITLE"]).StringCellValue,
                    });
                }
            }
            Console.Out.WriteLine(courses.Count + " courses loaded from xls");
            return courses;
        }
    }
}
