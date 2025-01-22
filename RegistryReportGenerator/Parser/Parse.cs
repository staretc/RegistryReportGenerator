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
            string id;
            string softwareName;

            // пока не дошли до конца
            while (!IsSheetEnded(++currentLine))
            {
                softwareName = _worksheet.Cell($"B{currentLine}").Value.ToString();

                // если есть гос. регистрация, добавляем запись о ней
                if (IsStateRegistered(currentLine))
                {
                    currentWriteDate = DateTime.Parse(_worksheet.Cell($"X{currentLine}").Value.ToString());
                    // если дата записи попадает в промежуток отбора, заносим запись
                    if (currentWriteDate >= beginningDate)
                    {
                        id = _worksheet.Cell($"W{currentLine}").Value.ToString();
                        eventType = "Гос. регистрация";
                        report.Add(new RegisrtyWrite(id, softwareName, currentWriteDate, eventType));
                    }
                }

                // если ячейка с датой исключения из реестра непустая, значит добавляем запись об исключении
                if (IsExcluded(currentLine))
                {
                    currentWriteDate = _worksheet.Cell($"F{currentLine}").GetDateTime();
                    // если дата записи попадает в промежуток отбора, заносим запись
                    if (currentWriteDate >= beginningDate)
                    {
                        id = _worksheet.Cell($"A{currentLine}").Value.ToString();
                        eventType = "Исключение из реестра";
                        report.Add(new RegisrtyWrite(id, softwareName, currentWriteDate, eventType));
                    }
                }

                // добавляем запись о включении в реестр

                currentWriteDate = _worksheet.Cell($"S{currentLine}").GetDateTime();
                // если дата записи попадает в промежуток отбора, заносим запись
                if (currentWriteDate >= beginningDate)
                {
                    id = _worksheet.Cell($"A{currentLine}").Value.ToString();
                    eventType = "Включение в реестр";
                    report.Add(new RegisrtyWrite(id, softwareName, currentWriteDate, eventType));
                }

                // добавляем запись о регистрации заявления
                currentWriteDate = _worksheet.Cell($"S{currentLine}").GetDateTime();
                // если дата записи попадает в промежуток отбора, заносим запись
                if (currentWriteDate >= beginningDate)
                {
                    id = _worksheet.Cell($"U{currentLine}").Value.ToString();
                    eventType = "Регистрация заявления";
                    report.Add(new RegisrtyWrite(id, softwareName, currentWriteDate, eventType));
                }
            }

            // возвращаем отсортированный по датам отчёт
            report.Sort();
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
            return _worksheet.Cell($"A{currentLine}").Value.ToString() == string.Empty;
        }
        /// <summary>
        /// Проверка, есть ли гос. регистрация текущей записи
        /// </summary>
        /// <param name="currentLine">Текущая проверяемая строка</param>
        /// <returns>Есть ли гос. регистрация</returns>
        private bool IsStateRegistered(int currentLine)
        {
            // дата гос. регистрации должна быть записана в столбце X
            return _worksheet.Cell($"X{currentLine}").Value.ToString() != string.Empty;
        }
    }
}
