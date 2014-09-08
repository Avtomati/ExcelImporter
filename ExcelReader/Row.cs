using System;
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
        public Cell GetCellAt(int index)
        {
            try
            {
                return Cells[index];
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public IList<Cell> Cells { get; private set; }
    }
}