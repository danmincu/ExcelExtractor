using System;

namespace ExcelExtractor2
{
    class Program
    {
        static int Main(string[] args)
        {
            // example of args
            // "c:\TACS\JETS VS PILOTS SKED 10-33-39.xlsx",T1%20Jets%20only;2;1;17;1;1;4|T2%20Jets%20only;2;1;5;1;1;2,2017-10-01,30,"30 DAYS JETS VS PILOTS SCHEDULE.xlsx"
            // output C:\Users\danmi\AppData\Local\Temp\30 DAYS JETS VS PILOTS SCHEDULE.xlsx
            try
            {
                Processor.Run(args);
                return 0;
            }
            catch
            {
                return 1;
            }
        }
    }
}
