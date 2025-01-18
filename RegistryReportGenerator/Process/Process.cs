using ClosedXML.Excel;
using Data;
using Parser;
using Source;
using Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Process
{
    public class ProcessGenerator
    {
        private XLWorkbook _workbook;
        private Report _report;

        /// <summary>
        /// Запускает работу программы
        /// </summary>
        /// <param name="source">Ссылка или путь к файлу, откуда берётся XLSX</param>
        /// <param name="outputFilePath">Путь к файлу, куда будет сохранён отчёт</param>
        public void Start(DateTime beginningDate, string source, string outputFilePath)
        {
            // отправляем сообщение о начале работы приложения
            Console.WriteLine("Начало работы . . .");

            // проверяем валидность выходного файла
            if (IsValidOutputPath(outputFilePath))
            {
                // проверяем валидность источника (должен быть ссылкой или путём к файлу xlsx)
                if (IsValidUrl(source))
                {
                    var dataCollecor = new GetDataFromUrl();
                    _workbook = dataCollecor.GetContent(source);
                }
                else if (IsValidInputPath(source))
                {
                    var dataCollecor = new GetDataFromFile();
                    _workbook = dataCollecor.GetContent(source);

                    // в случае использования файла другим процессом
                    if (_workbook == null)
                    {
                        return;
                    }
                }
                // если по источнику получили файл xlsx
                if (_workbook != null)
                {
                    // Отправляем уведомление об успешном открытии файла и начале работы с данными
                    Console.WriteLine("Файл xlsx успешно открыт!");
                    Console.WriteLine("Обработка данных . . .");

                    // ищем подходящие записи и сохраняем в _report
                    var converter = new Converter(_workbook.Worksheet("Данные"));
                    _report = converter.GenerateReport(beginningDate);

                    // сохраняем полученный отчёт
                    var store = new StoreReport();
                    store.SaveToFile(outputFilePath, _report);
                    Console.WriteLine("Отчёт подготовлен и сохранён успешно!");

                    return;
                }
                // если не получили xlsx из источника
                Console.WriteLine("Введён некорректный источник.");

                return;
            }
            // если выходной файл оказался не валидным
            Console.WriteLine("Введён некорректный путь к выходному файлу");
        }
        /// <summary>
        /// Проверяет валидность URL
        /// </summary>
        /// <param name="url">Ссылка, которую необходимо проверить</param>
        /// <returns>Результат проверки на валидность</returns>
        private bool IsValidUrl(string url)
        {
            // Ссылка должна быть полной и начинаться с http или https
            return Uri.TryCreate(url, UriKind.Absolute, out Uri uriResult) &&
                   (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
        }
        /// <summary>
        /// Проверяет валидность пути к входному файлу
        /// </summary>
        /// <param name="path">Путь, который необходимо проверить</param>
        /// <returns>Результат проверки на валидность</returns>
        private bool IsValidInputPath(string path)
        {
            // Путь не должен содержать недопустимые символы, должен быть абсолютьным и по данному пути должен существовать файл с разрешением xlsx
            return path.IndexOfAny(Path.GetInvalidPathChars()) == -1 &&
                   Path.IsPathRooted(path) &&
                   File.Exists(path) &&
                   string.Equals(Path.GetExtension(path), ".xlsx");
        }
        /// <summary>
        /// Проверяет валидность пути к выходному файлу
        /// </summary>
        /// <param name="path">Путь, который необходимо проверить</param>
        /// <returns>Результат проверки на валидность</returns>
        private bool IsValidOutputPath(string path)
        {
            // Путь не должен содержать недопустимые символы, должен быть абсолютьным и по данному пути должен существовать файл с разрешением xlsx
            return path.IndexOfAny(Path.GetInvalidPathChars()) == -1 &&
                   Directory.Exists(Path.GetDirectoryName(path));
        }
    }
}
