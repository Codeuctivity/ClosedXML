using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel
{
    /// <summary>
    /// 	Common methods
    /// </summary>
    public static partial class XLHelper
    {
        public const int MinRowNumber = 1;
        public const int MinColumnNumber = 1;
        public const int MaxRowNumber = 1048576;
        public const int MaxColumnNumber = 16384;
        public static readonly string MaxColumnLetter = "XFD";
        public static readonly double Epsilon = 1e-10;

        public static string LastCell => $"{MaxColumnLetter}{MaxRowNumber}";

        internal const NumberStyles NumberStyle = NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign | NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite | NumberStyles.AllowExponent;
        internal static readonly CultureInfo ParseCulture = CultureInfo.InvariantCulture;

        internal static readonly Regex RCSimpleRegex = new Regex(
            @"^(r(((-\d)?\d*)|\[(-\d)?\d*\]))?(c(((-\d)?\d*)|\[(-\d)?\d*\]))?$"
            , RegexOptions.IgnoreCase | RegexOptions.Compiled);

        internal static readonly Regex A1SimpleRegex = new Regex(
            @"\A"
            + @"(?<Reference>" // Start Group to pick
            + @"(?<Sheet>" // Start Sheet Name, optional
            + @"("
            + @"\'([^\[\]\*/\\\?:\']+|\'\')\'"
            // Sheet name with special characters, surrounding apostrophes are required
            + @"|"
            + @"\'?\w+\'?" // Sheet name with letters and numbers, surrounding apostrophes are optional
            + @")"
            + @"!)?" // End Sheet Name, optional
            + @"(?<Range>" // Start range
            + @"\$?[a-zA-Z]{1,3}\$?\d{1,7}" // A1 Address 1
            + @"(?<RangeEnd>:\$?[a-zA-Z]{1,3}\$?\d{1,7})?" // A1 Address 2, optional
            + @"|"
            + @"(?<ColumnNumbers>\$?\d{1,7}:\$?\d{1,7})" // 1:1
            + @"|"
            + @"(?<ColumnLetters>\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})" // A:A
            + @")" // End Range
            + @")" // End Group to pick
            + @"\Z"
            , RegexOptions.Compiled);

        internal static readonly Regex NamedRangeReferenceRegex =
            new Regex(@"^('?(?<Sheet>[^'!]+)'?!(?<Range>.+))|((?<Table>[^\[]+)\[(?<Column>[^\]]+)\])$",
                      RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.ExplicitCapture
                );

        private static readonly string[] letters = new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

        private static readonly string[] allLetters;
        private static readonly Dictionary<string, int> letterIndexes;

        static XLHelper()
        {
            allLetters = new string[MaxColumnNumber];
            letterIndexes = new Dictionary<string, int>(MaxColumnNumber, StringComparer.Create(ParseCulture, true));
            for (var i = 0; i < MaxColumnNumber; i++)
            {
                string letter;
                if (i < 26)
                {
                    letter = letters[i];
                }
                else if (i < 26 * 27)
                {
                    letter = letters[i / 26 - 1] + letters[i % 26];
                }
                else
                {
                    letter = letters[(i - 26) / 26 / 26 - 1] + letters[(i / 26 - 1) % 26] + letters[i % 26];
                }

                allLetters[i] = letter;
                letterIndexes.Add(letter, i + 1);
            }
        }

        /// <summary>
        /// Gets the column number of a given column letter.
        /// </summary>
        /// <param name="columnLetter"> The column letter to translate into a column number. </param>
        public static int GetColumnNumberFromLetter(string columnLetter)
        {
            if (string.IsNullOrEmpty(columnLetter))
            {
                throw new ArgumentNullException(nameof(columnLetter));
            }

            //Extra check because we allow users to pass row col positions in as strings
            if (columnLetter[0] <= '9')
            {
                return int.Parse(columnLetter, NumberStyle, ParseCulture);
            }

            if (letterIndexes.TryGetValue(columnLetter, out var retVal))
            {
                return retVal;
            }

            throw new ArgumentOutOfRangeException(columnLetter + " is not recognized as a column letter");
        }

        /// <summary>
        /// Gets the column letter of a given column number.
        /// </summary>
        /// <param name="columnNumber">The column number to translate into a column letter.</param>
        /// <param name="trimToAllowed">if set to <c>true</c> the column letter will be restricted to the allowed range.</param>
        /// <returns></returns>
        public static string GetColumnLetterFromNumber(int columnNumber, bool trimToAllowed = false)
        {
            if (trimToAllowed)
            {
                columnNumber = TrimColumnNumber(columnNumber);
            }

            if (columnNumber <= 0 || columnNumber > allLetters.Length)
            {
                throw new ArgumentOutOfRangeException(nameof(columnNumber));
            }

            // Adjust for start on column 1
            return allLetters[columnNumber - 1];
        }

        internal static int TrimColumnNumber(int columnNumber)
        {
            return Math.Max(MinColumnNumber, Math.Min(MaxColumnNumber, columnNumber));
        }

        internal static int TrimRowNumber(int rowNumber)
        {
            return Math.Max(MinRowNumber, Math.Min(MaxRowNumber, rowNumber));
        }

        public static bool IsValidColumn(string column)
        {
            if (string.IsNullOrWhiteSpace(column))
            {
                return false;
            }

            var length = column.Length;
            if (length > 3)
            {
                return false;
            }

            var theColumn = column.ToUpper();

            var isValid = theColumn[0] >= 'A' && theColumn[0] <= 'Z';
            if (length == 1)
            {
                return isValid;
            }

            if (length == 2)
            {
                return isValid && theColumn[1] >= 'A' && theColumn[1] <= 'Z';
            }

            if (theColumn[0] >= 'A' && theColumn[0] < 'X')
            {
                return theColumn[1] >= 'A' && theColumn[1] <= 'Z'
                       && theColumn[2] >= 'A' && theColumn[2] <= 'Z';
            }

            if (theColumn[0] != 'X')
            {
                return false;
            }

            if (theColumn[1] < 'F')
            {
                return theColumn[2] >= 'A' && theColumn[2] <= 'Z';
            }

            if (theColumn[1] != 'F')
            {
                return false;
            }

            return theColumn[2] >= 'A' && theColumn[2] <= 'D';
        }

        public static bool IsValidRow(string rowString)
        {
            if (int.TryParse(rowString, out var row))
            {
                return row > 0 && row <= MaxRowNumber;
            }

            return false;
        }

        public static bool IsValidA1Address(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                return false;
            }

            address = address.Replace("$", "");
            var rowPos = 0;
            var addressLength = address.Length;
            while (rowPos < addressLength && (address[rowPos] > '9' || address[rowPos] < '0'))
            {
                rowPos++;
            }

            return
                rowPos < addressLength
                && IsValidRow(address.Substring(rowPos))
                && IsValidColumn(address.Substring(0, rowPos));
        }

        public static bool IsValidRCAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                return false;
            }

            return RCSimpleRegex.IsMatch(address);
        }

        public static bool IsValidRangeAddress(string rangeAddress)
        {
            return A1SimpleRegex.IsMatch(rangeAddress);
        }

        public static bool IsValidRangeAddress(IXLRangeAddress rangeAddress)
        {
            return rangeAddress.IsValid
                   && rangeAddress.FirstAddress.RowNumber >= 1 && rangeAddress.LastAddress.RowNumber <= MaxRowNumber
                   && rangeAddress.FirstAddress.ColumnNumber >= 1 && rangeAddress.LastAddress.ColumnNumber <= MaxColumnNumber
                   && rangeAddress.FirstAddress.RowNumber <= rangeAddress.LastAddress.RowNumber
                   && rangeAddress.FirstAddress.ColumnNumber <= rangeAddress.LastAddress.ColumnNumber;
        }

        public static int GetColumnNumberFromAddress(string cellAddressString)
        {
            var rowPos = 0;
            while (cellAddressString[rowPos] > '9')
            {
                rowPos++;
            }

            return GetColumnNumberFromLetter(cellAddressString.Substring(0, rowPos));
        }

        internal static string[] SplitRange(string range)
        {
            return range.Contains('-') ? range.Replace('-', ':').Split(':') : range.Split(':');
        }

        public static int GetPtFromPx(double px)
        {
            return Convert.ToInt32(px * 72.0);
        }

        public static double GetPxFromPt(int pt)
        {
            return Convert.ToDouble(pt) / 72.0;
        }

        internal static IXLTableRows InsertRowsWithoutEvents(Func<int, bool, IXLRangeRows> insertFunc,
                                                             XLTableRange tableRange, int numberOfRows,
                                                             bool expandTable)
        {
            var ws = tableRange.Worksheet;
            var tracking = ws.EventTrackingEnabled;
            ws.EventTrackingEnabled = false;

            var rows = new XLTableRows(ws.Style);
            var inserted = insertFunc(numberOfRows, false);
            inserted.ForEach(r => rows.Add(new XLTableRow(tableRange, r as XLRangeRow)));

            if (expandTable)
            {
                tableRange.Table.ExpandTableRows(numberOfRows);
            }

            ws.EventTrackingEnabled = tracking;

            return rows;
        }

#if false
// Not using this anymore, but keeping it around for in case we bring back .NET3.5 support.
        public static bool IsNullOrWhiteSpace(string value)
        {
#if _NET35_
            if (value == null) return true;
            return value.All(c => char.IsWhiteSpace(c));
#else
            return String.IsNullOrWhiteSpace(value);
#endif
        }
#endif

        private static readonly Regex A1RegexRelative = new Regex(
      @"(?<=\W)(?<one>\$?[a-zA-Z]{1,3}\$?\d{1,7})(?=\W)" // A1
    + @"|(?<=\W)(?<two>\$?\d{1,7}:\$?\d{1,7})(?=\W)" // 1:1
    + @"|(?<=\W)(?<three>\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})(?=\W)", RegexOptions.Compiled); // A:A

        private static string Evaluator(Match match, int row, string column)
        {
            if (match.Groups["one"].Success)
            {
                var split = match.Groups["one"].Value.Split('$');
                if (split.Length == 1)
                {
                    return column + row; // A1
                }

                if (split.Length == 3)
                {
                    return match.Groups["one"].Value; // $A$1
                }

                var a = XLAddress.Create(match.Groups["one"].Value);
                if (split[0].Length == 0)
                {
                    return "$" + a.ColumnLetter + row; // $A1
                }

                return column + "$" + a.RowNumber;
            }

            if (match.Groups["two"].Success)
            {
                return ReplaceGroup(match.Groups["two"].Value, row.ToString());
            }

            return ReplaceGroup(match.Groups["three"].Value, column);
        }

        private static string ReplaceGroup(string value, string item)
        {
            var split = value.Split(':');
            var ret1 = split[0].StartsWith("$") ? split[0] : item;
            var ret2 = split[1].StartsWith("$") ? split[1] : item;
            return ret1 + ":" + ret2;
        }

        internal static string ReplaceRelative(string value, int row, string column)
        {
            var oldValue = ">" + value + "<";
            var newVal = A1RegexRelative.Replace(oldValue, m => Evaluator(m, row, column));
            return newVal.Substring(1, newVal.Length - 2);
        }

        public static bool AreEqual(double d1, double d2)
        {
            return Math.Abs(d1 - d2) < Epsilon;
        }

        public static DateTime GetDate(object v)
        {
            if (v is DateTime time)
            {
                return time;
            }

            if (v is double @double && @double.IsValidOADateNumber())
            {
                return DateTime.FromOADate(@double);
            }

            return (DateTime)Convert.ChangeType(v, typeof(DateTime));
        }

        internal static bool IsValidOADateNumber(this double d)
        {
            return -657435 <= d && d < 2958466;
        }

        /// <summary>
        /// A backward compatible version of <see cref="TimeSpan.FromDays(double)"/> that returns a value
        /// rounded to milliseconds. In .Net Core 3.0 the behavior has changed and timespan includes microseconds
        /// as well. As a result, the value 1:12:30 saved on one machine could become 1:12:29.999999 on another.
        /// </summary>
        internal static TimeSpan GetTimeSpan(double totalDays)
        {
            var timeSpan = TimeSpan.FromDays(totalDays);
            var roundedMilliseconds = Math.Round(timeSpan.TotalMilliseconds);
            return TimeSpan.FromMilliseconds(roundedMilliseconds);
        }

        public static bool ValidateName(string objectType, string newName, string oldName, IEnumerable<string> existingNames, out string message)
        {
            message = "";
            if (string.IsNullOrWhiteSpace(newName))
            {
                message = $"The {objectType} name '{newName}' is invalid";
                return false;
            }

            // Table names are case insensitive
            if (!oldName.Equals(newName, StringComparison.OrdinalIgnoreCase)
                && existingNames.Contains(newName, StringComparer.OrdinalIgnoreCase))
            {
                message = $"There is already a {objectType} named '{newName}'";
                return false;
            }

            var allowedFirstCharacters = new[] { '_', '\\' };
            if (!allowedFirstCharacters.Contains(newName[0]) && !char.IsLetter(newName[0]))
            {
                message = $"The {objectType} name '{newName}' does not begin with a letter, an underscore or a backslash.";
                return false;
            }

            if (newName.Length > 255)
            {
                message = $"The {objectType} name is more than 255 characters";
                return false;
            }

            if (new[] { 'C', 'R' }.Any(c => newName.ToUpper().Equals(c.ToString())))
            {
                message = $"The {objectType} name '{newName}' is invalid";
                return false;
            }

            return true;
        }
    }
}