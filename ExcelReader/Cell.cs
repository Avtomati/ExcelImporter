using System;

namespace ExcelReader
{
    public class Cell
    {
        readonly object _value;
        public Cell(object value)
        {
            _value = value;
        }

        public string AsString()
        {
            return (_value != null ? _value.ToString() : null);
        }

        public double? AsDouble()
        {
            if (_value == null) return null;
            if (_value is double) return (double)_value;
            if (_value is decimal) return (double)((decimal)_value);
            if (_value is int) return (int)_value;
            if (_value is long) return (long)_value; 
            double v;
            return (double.TryParse(_value.ToString(), out v) ? v : default(double?));
        }

        public decimal? AsDecimal()
        {
            if (_value == null) return null;
            if (_value is decimal) return (decimal)_value;
            if (_value is double) return (decimal)((double)_value);
            if (_value is int) return (int)_value;
            if (_value is long) return (long)_value;
            decimal v;
            return (decimal.TryParse(_value.ToString(), out v) ? v : default(decimal?));
        }
        public DateTime? AsDateTime()
        {
            if (_value == null) return null;
            if (_value is DateTime) return (DateTime)_value;
            DateTime v;
            return (DateTime.TryParse(_value.ToString(), out v) ? v : default(DateTime?));
        }

        public override string ToString()
        {
            return string.Format("{0}", _value);
        }

    }
}