using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelReader
{
    public class ExcelImporter
    {
        private readonly WorkBookReader[] _readers;

        public ExcelImporter(params WorkBookReader[] readers)
        {
            this._readers = readers;
        }

        public WorkBook Import(string filePath)
        {
            var reader = _readers.FirstOrDefault(r => r.CanRead(filePath));
            if (reader == null) throw new InvalidOperationException("can't read! do not have specific reader for file.");
            var sheets = GetSheets(reader.Read(filePath));
            return new WorkBook(filePath, sheets);
        }

        static IEnumerable<Sheet> GetSheets(Func<IEnumerable<Tuple<string, Func<IEnumerable<Tuple<int, int, object>>>>>> wbEnum)
        {
            return from sheet in wbEnum() 
                   let rows = CellsToRows(sheet.Item2()).ToList() 
                   where rows.Count > 0 
                   select new Sheet(sheet.Item1, rows);
        }

        static IEnumerable<Row> CellsToRows(IEnumerable<Tuple<int, int, object>> sheet)
        {

            var cells = new List<Cell>();
            var lastRowIndex = default(int?);
            var lastCol = 1;
            foreach (var cell in sheet)
            {
                if (lastRowIndex.HasValue && lastRowIndex.Value != cell.Item1)
                {
                    if (cells.Any(x => x != null))
                    {
                        yield return new Row(cells);
                    }
                    cells = new List<Cell>();
                    lastCol = 1;
                }
                lastRowIndex = cell.Item1;

                if (lastCol < cell.Item2)
                {
                    cells.AddRange(Enumerable.Range(lastCol, cell.Item2 - lastCol).Select(i => new Cell(null)));
                }
                lastCol = cell.Item2 + 1;

                cells.Add(new Cell(cell.Item3));
            }
            if (cells.Any(x => x != null))
            {
                yield return new Row(cells);
            }
        }
    }
}