using ClosedXML.Attributes;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ClosedXML.Examples.Tables
{
    public class InsertingTables : IXLExample
    {
        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Inserting Tables");

            // From a list of strings
            var listOfStrings = new List<string>
            {
                "House",
                "Car"
            };
            ws.Cell(1, 1).Value = "From Strings";
            ws.Cell(1, 1).AsRange().AddToNamed("Titles");
            ws.Cell(2, 1).InsertTable(listOfStrings);

            // From a list of arrays
            var listOfArr = new List<int[]>
            {
                new [] { 1, 2, 3 },
                new [] { 1 },
                new [] { 1, 2, 3, 4, 5, 6 }
            };
            ws.Cell(1, 3).Value = "From Arrays";
            ws.Range(1, 3, 1, 8).Merge().AddToNamed("Titles");
            ws.Cell(2, 3).InsertTable(listOfArr);

            // From a DataTable
            using var dataTable = GetTable();
            ws.Cell(7, 1).Value = "From DataTable";
            ws.Range(7, 1, 7, 4).Merge().AddToNamed("Titles");
            ws.Cell(8, 1).InsertTable(dataTable);

            // From a query
            var list = new List<Person>
            {
                new Person { Name = "John", Age = 30, House = "On Elm St." },
                new Person { Name = "Mary", Age = 15, House = "On Main St." },
                new Person { Name = "Luis", Age = 21, House = "On 23rd St." },
                new Person { Name = "Henry", Age = 45, House = "On 5th Ave." }
            };

            var people = from p in list
                         where p.Age >= 21
                         select p;

            ws.Cell(7, 6).Value = "From Query";
            ws.Range(7, 6, 7, 9).Merge().AddToNamed("Titles");
            ws.Cell(8, 6).InsertTable(people);

            ws.Cell(15, 6).Value = "From List";
            ws.Range(15, 6, 15, 9).Merge().AddToNamed("Titles");
            ws.Cell(16, 6).InsertTable(people);

            // Prepare the style for the titles
            var titlesStyle = wb.Style;
            titlesStyle.Font.Bold = true;
            titlesStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titlesStyle.Fill.BackgroundColor = XLColor.Cyan;

            // Format all titles in one shot
            wb.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle;

            ws.Columns().AdjustToContents();

            wb.SaveAs(filePath);
        }

        private class Person
        {
            [XLColumn(Header = "House Street")]
            public string House { get; set; }

            public string Name { get; set; }
            public int Age { get; set; }

            [XLColumn(Header = "Class Type")]
            public static string ClassType => nameof(Person);
        }

        // Private
        private DataTable GetTable()
        {
            var table = new DataTable();
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1));
            table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2));
            table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3));
            table.Rows.Add(21, "Combivent", "Janet", new DateTime(2000, 1, 4));
            table.Rows.Add(100, "Dilantin", "Melanie", new DateTime(2000, 1, 5));
            return table;
        }

        // Override

        #endregion Methods
    }
}