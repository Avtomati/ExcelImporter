using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelReader
{
    public class XlsxWorkBookReader : WorkBookReader
    {
        public override bool CanRead(string filePath)
        {
            return String.Equals(Path.GetExtension(filePath), ".xlsx", StringComparison.OrdinalIgnoreCase);
        }

        public override Func<IEnumerable<Tuple<string, Func<IEnumerable<Tuple<int, int, object>>>>>> Read(string filePath)
        {
            return () =>
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                using (var ep = new OfficeOpenXml.ExcelPackage(stream))
                {
                    return ep.Workbook.Worksheets.Select(sheet =>
                        Tuple.Create<string, Func<IEnumerable<Tuple<int, int, object>>>>(sheet.Name, () =>
                            sheet.Cells["a:z"].Select(c => Tuple.Create(c.Start.Row, c.Start.Column, c.Value))
                            )
                        );
                }
            };
        }
    }
}
