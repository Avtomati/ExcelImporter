using System;
using System.Globalization;

namespace ExcelReader
{
    public class Cell
    {
        readonly object _value;

        public Cell(object value)
        {
            _value = value;
        }

        public object AsObject()
        {
            return _value;
        }

        public string AsString()
        {
            return (_value != null ? _value.ToString() : null);
        }

        public int? AsInt()
        {
             if (_value == null) return null;
            if (_value is int) return (int)_value;
            if (_value is decimal) return (int)((decimal)_value);
            if (_value is double) return (int)((double)_value);
            if (_value is long) return (int)((long)_value); 
            int v;
            return (int.TryParse(_value.ToString(), out v) ? v : default(int?));
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
            DateTime result;
            var formats = new[] { "dd.MM.yyyy", "dd/MM/yyyy" };
            return DateTime.TryParseExact(_value.ToString(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out result)
                ? result : default(DateTime?);
        }

        public override string ToString()
        {
            return string.Format("{0}", _value);
        }

    }
}