using System;
using Import_Export_Excel.Models;

namespace Import_Export_Excel.Services.Interfaces
{
	public interface IExcelExportService
	{
        byte[] ExportPersonsToExcel(List<Person> persons);
        List<Person> ImportPersonsFromExcel(Stream fileStream);
    }
}

