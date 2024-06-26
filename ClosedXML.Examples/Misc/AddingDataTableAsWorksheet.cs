using ClosedXML.Excel;
using System;
using System.Data;
using System.Linq;

namespace ClosedXML.Examples.Misc
{
    public class AddingDataTableAsWorksheet : IXLExample
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

            using var dataTable = GetTable("Information");

            // Add a DataTable as a worksheet
            wb.Worksheets.Add(dataTable);
            wb.Worksheets.First().Columns().AdjustToContents();

            wb.SaveAs(filePath);
        }

        // Private
        private DataTable GetTable(string tableName)
        {
            var table = new DataTable
            {
                TableName = tableName
            };
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