using System.Collections.Generic;
using System.Linq;

namespace ExcelReader
{
    public class WorkBook
    {
        public WorkBook(string filePath, IEnumerable<Sheet> sheets)
        {
            FilePath = filePath;
            Sheets = sheets.ToList();
        }
        public string FilePath { get; private set; }
        public IList<Sheet> Sheets { get; private set; }

    }
}