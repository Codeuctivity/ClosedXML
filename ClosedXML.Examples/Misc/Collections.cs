using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ClosedXML.Examples.Misc
{
    public class Collections : IXLExample
    {
        #region Variables

        // Public

        // Private

        #endregion Variables

        #region Properties

        // Public

        // Private

        // Override

        #endregion Properties

        #region Events

        // Public

        // Private

        // Override

        #endregion Events

        #region Methods

        // Public
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Collections");

            // From a list of strings
            var listOfStrings = new List<string>
            {
                "House",
                "Car"
            };
            ws.Cell(1, 1).Value = "Strings";
            ws.Cell(1, 1).AsRange().AddToNamed("Titles");
            ws.Cell(2, 1).Value = listOfStrings;

            // From a list of arrays
            var listOfArr = new List<int[]>
            {
                new [] { 1, 2, 3 },
                new [] { 1 },
                new [] { 1, 2, 3, 4, 5, 6 }
            };
            ws.Cell(1, 3).Value = "Arrays";
            ws.Range(1, 3, 1, 8).Merge().AddToNamed("Titles");
            ws.Cell(2, 3).Value = listOfArr;

            // From a DataTable
            var dataTable = GetTable();
            ws.Cell(6, 1).Value = "DataTable";
            ws.Range(6, 1, 6, 4).Merge().AddToNamed("Titles");
            ws.Cell(7, 1).Value = dataTable;

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
                         select new { p.Name, p.House, p.Age };

            ws.Cell(6, 6).Value = "Query";
            ws.Range(6, 6, 6, 8).Merge().AddToNamed("Titles");
            ws.Cell(7, 6).Value = people;

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
            public string House { get; set; }
            public string Name { get; set; }
            public int Age { get; set; }
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