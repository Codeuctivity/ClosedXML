﻿
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.Worksheets
{
    [TestFixture]
    public class XLSheetViewTests
    {
        [Test]
        public void CopyWorksheetSheetViews()
        {
            using var wb1 = new XLWorkbook();
            using var wb2 = new XLWorkbook();

            var ws1 = wb1.AddWorksheet("WS1");
            ws1.SheetView.TopLeftCellAddress = ws1.Cell("AZ2000").Address;

            var ws2 = ws1.CopyTo(wb2, "WS2");

            Assert.That(ws2.SheetView.Worksheet, Is.EqualTo(ws2));
            Assert.That(ws2.SheetView.TopLeftCellAddress.ToString(), Is.EqualTo("AZ2000"));
        }

        [Test]
        public void InvalidTopLeftCell()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet();
            var ws2 = wb.AddWorksheet();

            Assert.Throws<ArgumentException>(() => ws1.SheetView.TopLeftCellAddress = ws2.Cell("A1").Address);
        }

        [Test]
        public void SheetViews()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                ws.SheetView.TopLeftCellAddress = ws.Cell("AZ2000").Address;
                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                Assert.That(ws.SheetView.TopLeftCellAddress.ToString(), Is.EqualTo("AZ2000"));

                ws.SheetView.TopLeftCellAddress = ws.Cell("AZ2000")
                    .CellBelow()
                    .CellRight()
                    .Address;

                wb.Save();
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                Assert.That(ws.SheetView.TopLeftCellAddress.ToString(), Is.EqualTo("BA2001"));
            }
        }
    }
}
