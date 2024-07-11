using System;
using Import_Export_Excel.Models;
using Import_Export_Excel.Services.Interfaces;
using OfficeOpenXml;

namespace Import_Export_Excel.Services
{
    public class ExcelExportService : IExcelExportService
    {
        public ExcelExportService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set the license context here
        }
        public byte[] ExportPersonsToExcel(List<Person> persons)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Persons");
                worksheet.Cells[1, 1].Value = "Id";
                worksheet.Cells[1, 2].Value = "Name";
                worksheet.Cells[1, 3].Value = "Surname";
                worksheet.Cells[1, 4].Value = "Age";
                worksheet.Cells[1, 5].Value = "Email";
                worksheet.Cells[1, 6].Value = "CreatedAt";

                for (int i = 0; i < persons.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = persons[i].Id;
                    worksheet.Cells[i + 2, 2].Value = persons[i].Name;
                    worksheet.Cells[i + 2, 3].Value = persons[i].Surname;
                    worksheet.Cells[i + 2, 4].Value = persons[i].Age;
                    worksheet.Cells[i + 2, 5].Value = persons[i].Email;
                    worksheet.Cells[i + 2, 6].Value = persons[i].CreatedAt;
                }

                return package.GetAsByteArray();
            }

        }
        public List<Person> ImportPersonsFromExcel(Stream fileStream)
        {
            var persons = new List<Person>();

            using (var package = new ExcelPackage(fileStream))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first worksheet
                var rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++) // Assuming the first row contains headers
                {
                    var person = new Person
                    {
                        Id = int.Parse(worksheet.Cells[row, 1].Text),
                        Name = worksheet.Cells[row, 2].Text,
                        Surname = worksheet.Cells[row, 3].Text,
                        Age = int.Parse(worksheet.Cells[row, 4].Text),
                        Email = worksheet.Cells[row, 5].Text,
                        CreatedAt = DateTime.FromOADate(double.Parse(worksheet.Cells[row, 6].Text)) // Convert Excel serial number to DateTime
                    };

                    persons.Add(person);
                }
            }

            return persons;
        }
    }

}

