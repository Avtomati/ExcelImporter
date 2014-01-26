using System.Collections.Generic;
using System.Linq;

namespace ExcelReader
{
    public class Row
    {
        public Row(IEnumerable<Cell> cells)
        {
            Cells = cells.ToList();
        }
        public IList<Cell> Cells { get; private set; }
    }
}