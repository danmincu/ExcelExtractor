using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtractor2
{
    public class SheetInstructions
    {
        public SheetInstructions(string sheetName, Tuple<Position, Position> airplaneNames, Position calendarStartPosition, DateTime minimDate, DateTime maximDate)
        {
            this.SheetName = sheetName;
            this.AirplaneNames = airplaneNames;
            this.CalendarStartPosition = calendarStartPosition;
            this.MinimDate = minimDate;
            this.MaximDate = maximDate;
        }

        public string SheetName { set; get; }
        public Tuple<Position, Position> AirplaneNames { set; get; }
        public Position CalendarStartPosition { set; get; }
        public DateTime MinimDate { set; get; }
        public DateTime MaximDate { set; get; }
    }
}
