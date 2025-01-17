using ClosedXML.Excel;
using Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Parser
{
    /// <summary>
    /// Класс получения данных для генерации отчёта
    /// </summary>
    public class Converter
    {
        private readonly IXLWorksheet _worksheet;

        /// <summary>
        /// Создание объекта класса Converter
        /// </summary>
        /// <param name="worksheet">Рабочий лист XLSX</param>
        public Converter(IXLWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        /// <summary>
        /// Генерация отчёта по XLSX страничке
        /// </summary>
        /// <param name="worksheet">Страница файла XLSX, из которой берутся данные</param>
        /// <param name="beginningDate">Начало периода, с которого берутся записи </param>
        /// <returns>Подготовленный отчёт</returns>
        public Report GenerateReport(DateTime beginningDate)
        {
            Report report = new Report(beginningDate);
            DateTime currentWriteDate;
            int currentLine = 6;
            string eventType;
            while ((currentWriteDate = _worksheet.Cell($"S{currentLine}").GetDateTime()) >= beginningDate)
            {
                // проверка, дошли ли до конца листа
                if (IsSheetEnded(currentLine))
                {
                    break;
                }

                eventType = "Включение в реестр";
                // если ячейка с датой исключения из реестра непустая, значит запись об исключении
                if (IsExcluded(currentLine))
                {
                    currentWriteDate = _worksheet.Cell($"F{currentLine}").GetDateTime();
                    eventType = "Исключение из реестра";
                }
                // добавляем новую запись
                report.Add(new RegisrtyWrite(_worksheet.Cell($"A{currentLine}").Value.ToString(),
                                             _worksheet.Cell($"B{currentLine}").Value.ToString(),
                                             currentWriteDate,
                                             eventType));
                currentLine++;
            }

            return report;
        }
        /// <summary>
        /// Проверка на исключение из реестра
        /// </summary>
        /// <param name="currentLine">Текущая строка, в которой находится запись</param>
        /// <returns>Является ли запись исключённой из реестра</returns>
        private bool IsExcluded(int currentLine)
        {
            // дата исключения из реестра должна быть записана в столбце F
            return _worksheet.Cell($"F{currentLine}").Value.ToString() != string.Empty;
        }
        /// <summary>
        /// Проверка, дошли ли до конца листа
        /// </summary>
        /// <param name="currentLine">Текущая проверяемая строка</param>
        /// <returns>Дошли ли до конца</returns>
        private bool IsSheetEnded(int currentLine)
        {
            // прове
            return _worksheet.Cell($"A{currentLine}").Value.ToString() == string.Empty;
        }
    }
}
