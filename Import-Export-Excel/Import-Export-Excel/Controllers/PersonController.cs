using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Import_Export_Excel.ExcelFiles;
using Import_Export_Excel.Models;
using Import_Export_Excel.Services.Interfaces;
using Microsoft.AspNetCore.Mvc;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace Import_Export_Excel.Controllers
{
    [Route("api/[controller]")]
    public class PersonController : Controller
    {
        private readonly IExcelExportService _excelExportService;

        public PersonController(IExcelExportService excelExportService)
        {
            _excelExportService = excelExportService;
        }
        [HttpGet("export")]
        public IActionResult ExportPersons()
        {
            var persons = new List<Person>
        {
            new Person { Id = 1, Name = "John", Surname = "Doe", Age = 30, Email = "john.doe@example.com", CreatedAt = DateTime.Now },
            new Person { Id = 2, Name = "Rufat", Surname = "Ismayilov", Age = 35, Email = "Rufat.doe@example.com", CreatedAt = DateTime.Now },
            new Person { Id = 3, Name = "Asim", Surname = "Azizov", Age = 30, Email = "Asim.doe@example.com", CreatedAt = DateTime.Now },
            new Person { Id = 4, Name = "Qismat", Surname = "Macidov", Age = 40, Email = "Qismat.doe@example.com", CreatedAt = DateTime.Now },
            // Add more sample persons or fetch from your database
        };

            var fileContents = _excelExportService.ExportPersonsToExcel(persons);

            return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Persons.xlsx");
        }
        [HttpPost("import")]
        public IActionResult ImportPersons([FromForm] PersonTableDto personTableDto)
        {
            if (personTableDto.File == null || personTableDto.File.Length == 0)
                return BadRequest("Please upload a valid Excel file.");

            List<Person> persons;
            using (var stream = new MemoryStream())
            {
                personTableDto.File.CopyTo(stream);
                stream.Position = 0;
                persons = _excelExportService.ImportPersonsFromExcel(stream);
            }

            // Here you can save persons to the database or perform further processing
            return Ok(persons);
        }
       
    }
}




