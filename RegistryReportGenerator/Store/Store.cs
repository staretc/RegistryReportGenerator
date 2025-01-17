using Data;
using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Store
{
    public class StoreReport
    {
        /// <summary>
        /// Сохранение отчёта в файл
        /// </summary>
        /// <param name="filePath">Путь, по которому сохраняется файл</param>
        /// <param name="report">Подготовленный для сохранения отчёт</param>
        public void SaveToFile(string filePath, Report report)
        {
            File.WriteAllText(filePath, report.ToString());
        }
    }
}
