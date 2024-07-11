using System;
namespace Import_Export_Excel.Models
{
	public class Person
	{
        public int Id { get; set; }
        public string Name { get; set; }
        public string Surname { get; set; }
        public int Age { get; set; }
        public string Email { get; set; }
        public DateTime CreatedAt { get; set; }
        public Person()
		{
		}
	}
}

