using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.RichText
{
    /// <summary>
    ///     This is a test class for XLRichStringTests and is intended
    ///     to contain all XLRichStringTests Unit Tests
    /// </summary>
    [TestFixture]
    public class XLRichStringTests
    {
        [Test]
        public void AccessRichTextTest1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.CreateRichText().AddText("12");
            cell.DataType = XLDataType.Number;

            Assert.That(cell.GetDouble(), Is.EqualTo(12.0));

            var richText = cell.GetRichText();

            Assert.That(richText.ToString(), Is.EqualTo("12"));

            richText.AddText("34");

            Assert.That(cell.GetString(), Is.EqualTo("1234"));

            Assert.That(cell.DataType, Is.EqualTo(XLDataType.Number));

            Assert.That(cell.GetDouble(), Is.EqualTo(1234.0));
        }

        /// <summary>
        ///     A test for AddText
        /// </summary>
        [Test]
        public void AddTextTest1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var richString = cell.CreateRichText();

            var text = "Hello";
            richString.AddText(text).SetBold().SetFontColor(XLColor.Red);

            Assert.That(cell.GetString(), Is.EqualTo(text));
            Assert.That(cell.GetRichText().First().Bold, Is.True);
            Assert.That(XLColor.Red, Is.EqualTo(cell.GetRichText().First().FontColor));

            Assert.That(richString, Has.Count.EqualTo(1));

            richString.AddText("World");
            Assert.That(text, Is.EqualTo(richString.First().Text), "Item in collection is not the same as the one returned");
        }

        [Test]
        public void AddTextTest2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var number = 123;

            cell.SetValue(number).Style
                .Font.SetBold()
                .Font.SetFontColor(XLColor.Red);

            var text = number.ToString();

            Assert.That(text, Is.EqualTo(cell.GetRichText().ToString()));
            Assert.That(cell.GetRichText().First().Bold, Is.True);
            Assert.That(XLColor.Red, Is.EqualTo(cell.GetRichText().First().FontColor));

            Assert.That(cell.GetRichText(), Has.Count.EqualTo(1));

            cell.GetRichText().AddText("World");
            Assert.That(text, Is.EqualTo(cell.GetRichText().First().Text), "Item in collection is not the same as the one returned");
        }

        [Test]
        public void AddTextTest3()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var number = 123;
            cell.Value = number;
            cell.Style
                .Font.SetBold()
                .Font.SetFontColor(XLColor.Red);

            var text = number.ToString();

            Assert.That(text, Is.EqualTo(cell.GetRichText().ToString()));
            Assert.That(cell.GetRichText().First().Bold, Is.True);
            Assert.That(XLColor.Red, Is.EqualTo(cell.GetRichText().First().FontColor));

            Assert.That(cell.GetRichText().Count, Is.EqualTo(1));

            cell.GetRichText().AddText("World");
            Assert.That(text, Is.EqualTo(cell.GetRichText().First().Text), "Item in collection is not the same as the one returned");
        }

        /// <summary>
        ///     A test for Clear
        /// </summary>
        [Test]
        public void ClearTest()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World!");

            richString.ClearText();
            var expected = string.Empty;
            var actual = richString.ToString();
            Assert.That(actual, Is.EqualTo(expected));

            Assert.That(richString.Count, Is.EqualTo(0));
        }

        [Test]
        public void CountTest()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World!");

            Assert.That(richString.Count, Is.EqualTo(3));
        }

        [Test]
        public void HasRichTextTest1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.GetRichText().AddText("123");

            Assert.That(cell.HasRichText, Is.EqualTo(true));

            cell.DataType = XLDataType.Text;

            Assert.That(cell.HasRichText, Is.EqualTo(true));

            cell.DataType = XLDataType.Number;

            Assert.That(cell.HasRichText, Is.EqualTo(false));

            cell.GetRichText().AddText("123");

            Assert.That(cell.HasRichText, Is.EqualTo(true));

            cell.Value = 123;

            Assert.That(cell.HasRichText, Is.EqualTo(false));

            cell.GetRichText().AddText("123");

            Assert.That(cell.HasRichText, Is.EqualTo(true));

            cell.SetValue("123");

            Assert.That(cell.HasRichText, Is.EqualTo(false));
        }

        /// <summary>
        ///     A test for Characters
        /// </summary>
        [Test]
        public void Substring_All_From_OneString()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            var actual = richString.Substring(0);

            Assert.That(actual.First(), Is.EqualTo(richString.First()));

            Assert.That(actual.Count, Is.EqualTo(1));

            actual.First().SetBold();

            Assert.That(ws.Cell(1, 1).GetRichText().First().Bold, Is.EqualTo(true));
        }

        [Test]
        public void Substring_All_From_ThreeStrings()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(0);

            Assert.That(actual.ElementAt(0), Is.EqualTo(richString.ElementAt(0)));
            Assert.That(actual.ElementAt(1), Is.EqualTo(richString.ElementAt(1)));
            Assert.That(actual.ElementAt(2), Is.EqualTo(richString.ElementAt(2)));

            Assert.That(actual.Count, Is.EqualTo(3));
            Assert.That(richString.Count, Is.EqualTo(3));

            actual.First().SetBold();

            Assert.That(ws.Cell(1, 1).GetRichText().First().Bold, Is.EqualTo(true));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().Last().Bold, Is.EqualTo(false));
        }

        [Test]
        public void Substring_From_OneString_End()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            var actual = richString.Substring(2);

            Assert.That(actual.Count, Is.EqualTo(1)); // substring was in one piece

            Assert.That(richString.Count, Is.EqualTo(2)); // The text was split because of the substring

            Assert.That(actual.First().Text, Is.EqualTo("llo"));

            Assert.That(richString.First().Text, Is.EqualTo("He"));
            Assert.That(richString.Last().Text, Is.EqualTo("llo"));

            actual.First().SetBold();

            Assert.That(ws.Cell(1, 1).GetRichText().First().Bold, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().Last().Bold, Is.EqualTo(true));

            richString.Last().SetItalic();

            Assert.That(ws.Cell(1, 1).GetRichText().First().Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().Last().Italic, Is.EqualTo(true));

            Assert.That(actual.First().Italic, Is.EqualTo(true));

            richString.SetFontSize(20);

            Assert.That(ws.Cell(1, 1).GetRichText().First().FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().Last().FontSize, Is.EqualTo(20));

            Assert.That(actual.First().FontSize, Is.EqualTo(20));
        }

        [Test]
        public void Substring_From_OneString_Middle()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            var actual = richString.Substring(2, 2);

            Assert.That(actual.Count, Is.EqualTo(1)); // substring was in one piece

            Assert.That(richString.Count, Is.EqualTo(3)); // The text was split because of the substring

            Assert.That(actual.First().Text, Is.EqualTo("ll"));

            Assert.That(richString.First().Text, Is.EqualTo("He"));
            Assert.That(richString.ElementAt(1).Text, Is.EqualTo("ll"));
            Assert.That(richString.Last().Text, Is.EqualTo("o"));

            actual.First().SetBold();

            Assert.That(ws.Cell(1, 1).GetRichText().First().Bold, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold, Is.EqualTo(true));
            Assert.That(ws.Cell(1, 1).GetRichText().Last().Bold, Is.EqualTo(false));

            richString.Last().SetItalic();

            Assert.That(ws.Cell(1, 1).GetRichText().First().Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().Last().Italic, Is.EqualTo(true));

            Assert.That(actual.First().Italic, Is.EqualTo(false));

            richString.SetFontSize(20);

            Assert.That(ws.Cell(1, 1).GetRichText().First().FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().Last().FontSize, Is.EqualTo(20));

            Assert.That(actual.First().FontSize, Is.EqualTo(20));
        }

        [Test]
        public void Substring_From_OneString_Start()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            var actual = richString.Substring(0, 2);

            Assert.That(actual.Count, Is.EqualTo(1)); // substring was in one piece

            Assert.That(richString.Count, Is.EqualTo(2)); // The text was split because of the substring

            Assert.That(actual.First().Text, Is.EqualTo("He"));

            Assert.That(richString.First().Text, Is.EqualTo("He"));
            Assert.That(richString.Last().Text, Is.EqualTo("llo"));

            actual.First().SetBold();

            Assert.That(ws.Cell(1, 1).GetRichText().First().Bold, Is.EqualTo(true));
            Assert.That(ws.Cell(1, 1).GetRichText().Last().Bold, Is.EqualTo(false));

            richString.Last().SetItalic();

            Assert.That(ws.Cell(1, 1).GetRichText().First().Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().Last().Italic, Is.EqualTo(true));

            Assert.That(actual.First().Italic, Is.EqualTo(false));

            richString.SetFontSize(20);

            Assert.That(ws.Cell(1, 1).GetRichText().First().FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().Last().FontSize, Is.EqualTo(20));

            Assert.That(actual.First().FontSize, Is.EqualTo(20));
        }

        [Test]
        public void Substring_From_ThreeStrings_End1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(21);

            Assert.That(actual.Count, Is.EqualTo(1)); // substring was in one piece

            Assert.That(richString.Count, Is.EqualTo(4)); // The text was split because of the substring

            Assert.That(actual.First().Text, Is.EqualTo("bors!"));

            Assert.That(richString.ElementAt(0).Text, Is.EqualTo("Good Morning"));
            Assert.That(richString.ElementAt(1).Text, Is.EqualTo(" my "));
            Assert.That(richString.ElementAt(2).Text, Is.EqualTo("neigh"));
            Assert.That(richString.ElementAt(3).Text, Is.EqualTo("bors!"));

            actual.First().SetBold();

            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Bold, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Bold, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Bold, Is.EqualTo(true));

            richString.Last().SetItalic();

            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Italic, Is.EqualTo(true));

            Assert.That(actual.First().Italic, Is.EqualTo(true));

            richString.SetFontSize(20);

            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize, Is.EqualTo(20));

            Assert.That(actual.First().FontSize, Is.EqualTo(20));
        }

        [Test]
        public void Substring_From_ThreeStrings_End2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(13);

            Assert.That(actual.Count, Is.EqualTo(2));

            Assert.That(richString.Count, Is.EqualTo(4)); // The text was split because of the substring

            Assert.That(actual.ElementAt(0).Text, Is.EqualTo("my "));
            Assert.That(actual.ElementAt(1).Text, Is.EqualTo("neighbors!"));

            Assert.That(richString.ElementAt(0).Text, Is.EqualTo("Good Morning"));
            Assert.That(richString.ElementAt(1).Text, Is.EqualTo(" "));
            Assert.That(richString.ElementAt(2).Text, Is.EqualTo("my "));
            Assert.That(richString.ElementAt(3).Text, Is.EqualTo("neighbors!"));

            actual.ElementAt(1).SetBold();

            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Bold, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Bold, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Bold, Is.EqualTo(true));

            richString.Last().SetItalic();

            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Italic, Is.EqualTo(true));

            Assert.That(actual.ElementAt(0).Italic, Is.EqualTo(false));
            Assert.That(actual.ElementAt(1).Italic, Is.EqualTo(true));

            richString.SetFontSize(20);

            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize, Is.EqualTo(20));

            Assert.That(actual.ElementAt(0).FontSize, Is.EqualTo(20));
            Assert.That(actual.ElementAt(1).FontSize, Is.EqualTo(20));
        }

        [Test]
        public void Substring_From_ThreeStrings_Mid1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(5, 10);

            Assert.That(actual.Count, Is.EqualTo(2));

            Assert.That(richString.Count, Is.EqualTo(5)); // The text was split because of the substring

            Assert.That(actual.ElementAt(0).Text, Is.EqualTo("Morning"));
            Assert.That(actual.ElementAt(1).Text, Is.EqualTo(" my"));

            Assert.That(richString.ElementAt(0).Text, Is.EqualTo("Good "));
            Assert.That(richString.ElementAt(1).Text, Is.EqualTo("Morning"));
            Assert.That(richString.ElementAt(2).Text, Is.EqualTo(" my"));
            Assert.That(richString.ElementAt(3).Text, Is.EqualTo(" "));
            Assert.That(richString.ElementAt(4).Text, Is.EqualTo("neighbors!"));
        }

        [Test]
        public void Substring_From_ThreeStrings_Mid2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(5, 15);

            Assert.That(actual.Count, Is.EqualTo(3));

            Assert.That(richString.Count, Is.EqualTo(5)); // The text was split because of the substring

            Assert.That(actual.ElementAt(0).Text, Is.EqualTo("Morning"));
            Assert.That(actual.ElementAt(1).Text, Is.EqualTo(" my "));
            Assert.That(actual.ElementAt(2).Text, Is.EqualTo("neig"));

            Assert.That(richString.ElementAt(0).Text, Is.EqualTo("Good "));
            Assert.That(richString.ElementAt(1).Text, Is.EqualTo("Morning"));
            Assert.That(richString.ElementAt(2).Text, Is.EqualTo(" my "));
            Assert.That(richString.ElementAt(3).Text, Is.EqualTo("neig"));
            Assert.That(richString.ElementAt(4).Text, Is.EqualTo("hbors!"));
        }

        [Test]
        public void Substring_From_ThreeStrings_Start1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(0, 4);

            Assert.That(actual.Count, Is.EqualTo(1)); // substring was in one piece

            Assert.That(richString.Count, Is.EqualTo(4)); // The text was split because of the substring

            Assert.That(actual.First().Text, Is.EqualTo("Good"));

            Assert.That(richString.ElementAt(0).Text, Is.EqualTo("Good"));
            Assert.That(richString.ElementAt(1).Text, Is.EqualTo(" Morning"));
            Assert.That(richString.ElementAt(2).Text, Is.EqualTo(" my "));
            Assert.That(richString.ElementAt(3).Text, Is.EqualTo("neighbors!"));

            actual.First().SetBold();

            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Bold, Is.EqualTo(true));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Bold, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Bold, Is.EqualTo(false));

            richString.First().SetItalic();

            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Italic, Is.EqualTo(true));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Italic, Is.EqualTo(false));

            Assert.That(actual.First().Italic, Is.EqualTo(true));

            richString.SetFontSize(20);

            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize, Is.EqualTo(20));

            Assert.That(actual.First().FontSize, Is.EqualTo(20));
        }

        [Test]
        public void Substring_From_ThreeStrings_Start2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Good Morning");
            richString.AddText(" my ");
            richString.AddText("neighbors!");

            var actual = richString.Substring(0, 15);

            Assert.That(actual.Count, Is.EqualTo(2));

            Assert.That(richString.Count, Is.EqualTo(4)); // The text was split because of the substring

            Assert.That(actual.ElementAt(0).Text, Is.EqualTo("Good Morning"));
            Assert.That(actual.ElementAt(1).Text, Is.EqualTo(" my"));

            Assert.That(richString.ElementAt(0).Text, Is.EqualTo("Good Morning"));
            Assert.That(richString.ElementAt(1).Text, Is.EqualTo(" my"));
            Assert.That(richString.ElementAt(2).Text, Is.EqualTo(" "));
            Assert.That(richString.ElementAt(3).Text, Is.EqualTo("neighbors!"));

            actual.ElementAt(1).SetBold();

            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Bold, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold, Is.EqualTo(true));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Bold, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Bold, Is.EqualTo(false));

            richString.First().SetItalic();

            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Italic, Is.EqualTo(true));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Italic, Is.EqualTo(false));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Italic, Is.EqualTo(false));

            Assert.That(actual.ElementAt(0).Italic, Is.EqualTo(true));
            Assert.That(actual.ElementAt(1).Italic, Is.EqualTo(false));

            richString.SetFontSize(20);

            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize, Is.EqualTo(20));
            Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize, Is.EqualTo(20));

            Assert.That(actual.ElementAt(0).FontSize, Is.EqualTo(20));
            Assert.That(actual.ElementAt(1).FontSize, Is.EqualTo(20));
        }

        [Test]
        public void Substring_IndexOutsideRange1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            Assert.That(() => richString.Substring(50), Throws.TypeOf<IndexOutOfRangeException>());
        }

        [Test]
        public void Substring_IndexOutsideRange2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText("World");

            Assert.That(() => richString.Substring(50), Throws.TypeOf<IndexOutOfRangeException>());
        }

        [Test]
        public void Substring_IndexOutsideRange3()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");

            Assert.That(() => richString.Substring(1, 10), Throws.TypeOf<IndexOutOfRangeException>());
        }

        [Test]
        public void Substring_IndexOutsideRange4()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText("World");

            Assert.That(() => richString.Substring(5, 20), Throws.TypeOf<IndexOutOfRangeException>());
        }

        /// <summary>
        ///     A test for ToString
        /// </summary>
        [Test]
        public void ToStringTest()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var richString = ws.Cell(1, 1).GetRichText();

            richString.AddText("Hello");
            richString.AddText(" ");
            richString.AddText("World");
            var expected = "Hello World";
            var actual = richString.ToString();
            Assert.That(actual, Is.EqualTo(expected));

            richString.AddText("!");
            expected = "Hello World!";
            actual = richString.ToString();
            Assert.That(actual, Is.EqualTo(expected));

            richString.ClearText();
            expected = string.Empty;
            actual = richString.ToString();
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test(Description = "See #1361")]
        public void CanClearInlinedRichText()
        {
            using var outputStream = new MemoryStream();
            using (var inputStream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\InlinedRichText\ChangeRichText\inputfile.xlsx")))
            using (var workbook = new XLWorkbook(inputStream))
            {
                workbook.Worksheets.First().Cell("A1").Value = "";
                workbook.SaveAs(outputStream);
            }

            using var wb = new XLWorkbook(outputStream);
            Assert.That(wb.Worksheets.First().Cell("A1").Value, Is.EqualTo(""));
        }

        [Test]
        public void CanChangeInlinedRichText()
        {
            static void testRichText(IXLRichText richText)
            {
                Assert.That(richText, Is.Not.Null);
                Assert.That(richText.Any(), Is.True);
                Assert.That(richText.ElementAt(2).Text, Is.EqualTo("3"));
                Assert.That(richText.ElementAt(2).FontColor, Is.EqualTo(XLColor.Red));
            }

            using var outputStream = new MemoryStream();
            using (var inputStream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\InlinedRichText\ChangeRichText\inputfile.xlsx")))
            using (var workbook = new XLWorkbook(inputStream))
            {
                var richText = workbook.Worksheets.First().Cell("A1").GetRichText();
                testRichText(richText);
                richText.AddText(" - changed");
                workbook.SaveAs(outputStream);
            }

            using var wb = new XLWorkbook(outputStream);
            var cell = wb.Worksheets.First().Cell("A1");
            Assert.That(cell.ShareString, Is.False);
            Assert.That(cell.HasRichText, Is.True);
            var rt = cell.GetRichText();
            Assert.That(rt.ToString(), Is.EqualTo("Year (range: 3 yrs) - changed"));
            testRichText(rt);
        }

        [Test]
        public void ClearInlineRichTextWhenRelevant()
        {
            var expectedFilePath = @"Other\InlinedRichText\ChangeRichTextToFormula\output.xlsx";

            using var ms = new MemoryStream();
            TestHelper.CreateAndCompare(() =>
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet();
                    var cell = ws.FirstCell();

                    cell.GetRichText().AddText("Bold").SetBold().AddText(" and red").SetBold().SetFontColor(XLColor.Red);
                    cell.ShareString = false;

                    //wb.SaveAs(ms);
                    wb.SaveAs(ms);
                }
                ms.Seek(0, SeekOrigin.Begin);

                var wb2 = new XLWorkbook(ms);
                {
                    var ws = wb2.Worksheets.First();
                    var cell = ws.FirstCell();

                    cell.FormulaA1 = "=1 + 2";
                    wb2.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                //var expectedFileInVsSolution = Path.GetFullPath(Path.Combine("../../../", "Resource", expectedFilePath));
                //File.WriteAllBytes(expectedFileInVsSolution, ms.ToArray());

                return wb2;
            }, expectedFilePath);
        }
    }
}