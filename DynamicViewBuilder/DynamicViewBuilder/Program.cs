using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DynamicViewBuilder
{
    class Program
    {
        static void Main(string[] args)
        {
            StringBuilder sb = new StringBuilder();
            string template1 = System.IO.File.ReadAllText("ViewColumnTemplate1.txt");
           // template = " {0} {0} {1} {0}";
            for (int i = 0; i <= ((int)'P') - 65; i++)
            {
                var oneColumn = template1.Replace("{0}", ((char)(i + 65)).ToString()).Replace("{1}", (i + 1).ToString());
                sb.AppendLine(oneColumn);
            }
            System.IO.File.WriteAllText("all1.txt", sb.ToString());

            sb = new StringBuilder();
            string template2 = System.IO.File.ReadAllText("ViewColumnTemplate2.txt");
            // template = " {0} {0} {1} {0}";
            for (int i = 0; i <= ((int)'D') - 65; i++)
            {
                var oneColumn = template2.Replace("{0}", ((char)(i + 65)).ToString()).Replace("{1}", (i + 1).ToString());
                sb.AppendLine(oneColumn);
            }
            System.IO.File.WriteAllText("all2.txt", sb.ToString());
        }
    }
}
