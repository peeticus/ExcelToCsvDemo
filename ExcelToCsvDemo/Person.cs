using System;

using CsvHelper.Configuration.Attributes;

namespace ExcelToCsvDemo
{
    internal class Person
    {
        [Name("Id")]
        public int Id { get; set; }
        [Name("Name")]
        public string Name { get; set; }
        [Name("Occupation")]
        public string Occupation { get; set; }
        [Name("Date Entered")]
        public DateTime DateEntered { get; set; }
    }
}
