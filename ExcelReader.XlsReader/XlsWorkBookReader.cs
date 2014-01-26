using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelReader
{
    public class XlsWorkBookReader : WorkBookReader
    {
        public override bool CanRead(string filePath)
        {
            return String.Equals(Path.GetExtension(filePath), ".xls", StringComparison.OrdinalIgnoreCase);
        }

        public override Func<IEnumerable<Tuple<string, Func<IEnumerable<Tuple<int, int, object>>>>>> Read(string filePath)
        {
            return () =>
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    return ExcelLibrary.SpreadSheet.Workbook.Load(stream).Worksheets.Select(sheet =>
                        Tuple.Create<string, Func<IEnumerable<Tuple<int, int, object>>>>(sheet.Name, () =>
                            ToEnumerable(sheet.Cells.GetEnumerator()).Select(c =>
                            {
                                var cell = c.Right;
                                var value = default(object);
                                if (cell.Format.FormatType == ExcelLibrary.SpreadSheet.CellFormatType.Date
                                    || cell.Format.FormatType == ExcelLibrary.SpreadSheet.CellFormatType.DateTime
                                    || cell.Format.FormatType == ExcelLibrary.SpreadSheet.CellFormatType.Time
                                    ||
                                    (cell.Format.FormatType == ExcelLibrary.SpreadSheet.CellFormatType.Custom &&
                                     (cell.FormatString.Contains("yy") || cell.FormatString.Contains("mm")))
                                    )
                                {
                                    value = cell.DateTimeValue;
                                }
                                else
                                {
                                    value = cell.Value;
                                }
                                return Tuple.Create(c.Left.Left + 1, c.Left.Right + 1, value);
                            }
                                )
                            )
                        );
                }
            };
        }

        static IEnumerable<T> ToEnumerable<T>(IEnumerator<T> enumerator)
        {
            while (enumerator.MoveNext())
            {
                yield return enumerator.Current;
            }
        }
    }
}
