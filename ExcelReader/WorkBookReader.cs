using System;
using System.Collections.Generic;

namespace ExcelReader
{
    public abstract class WorkBookReader
    {
        protected WorkBookReader(){}
        public abstract bool CanRead(string filePath);
        public abstract Func<IEnumerable<Tuple<string, Func<IEnumerable<Tuple<int, int, object>>>>>> Read(string filePath);
    }
}