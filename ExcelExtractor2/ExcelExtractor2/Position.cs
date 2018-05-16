using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtractor2
{
    public class Position
    {
        public int RowIndex { set; get; }
        public int ColumnIndex { set; get; }
        public override string ToString()
        {
            return $"({this.ColumnIndex},{this.RowIndex})";
        }
    }

  
}
