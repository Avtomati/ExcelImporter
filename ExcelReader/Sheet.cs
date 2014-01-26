using System.Collections.Generic;
using System.Linq;

namespace ExcelReader
{
    public class Sheet
    {
        public Sheet(string name, IEnumerable<Row> rows)
        {
            Name = name;
            Rows = rows.ToList();
        }
        public string Name { get; private set; }
        public IList<Row> Rows { get; private set; }
    }
}