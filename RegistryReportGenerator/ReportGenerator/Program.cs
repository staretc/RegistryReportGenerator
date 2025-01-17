using Process;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            int[] splittedDate = args[0].Split('.').Select(x => int.Parse(x)).ToArray();
            DateTime beginningDate = new DateTime(splittedDate[2], splittedDate[1], splittedDate[0]);

            var process = new ProcessGenerator();
            // дата начала, источник, выходной файл
            process.Start(beginningDate, args[1], args[2]);

            //var process = new ProcessGenerator();
            //process.Start(new DateTime(2019, 12, 28),
            //              "https://reestr.digital.gov.ru/reestr/",
            //              "C:\\Users\\89201\\Downloads\\result.txt");
        }
    }
}
