using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace HtmlReportReestr
{
    #region DATA
    /// <summary>
    /// Класс, хранящий необходимую информацию из записи реестра
    /// </summary>
    public class RegisrtyWrite
    {
        #region Fields
        private string _id;
        private string _softwareName;
        private DateTime _eventDate;
        private string _eventType;
        #endregion

        #region Properties
        public string Id
        {
            get => _id;
            set => _id = value;
        }
        public string SofwareName
        {
            get => _softwareName;
            set => _softwareName = value;
        }
        public DateTime EventDate
        {
            get => _eventDate;
            set
            {
                _eventDate = value;
            }
        }
        public string EventType
        {
            get => _eventType;
            set => _eventType = value;
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Создание объекта класса RegistryWrite
        /// </summary>
        /// <param name="id">Номер реестровой записи</param>
        /// <param name="softwareName">Наименование ПО</param>
        /// <param name="date">Дата регистрации</param>
        /// <param name="eventType">Тип события</param>
        public RegisrtyWrite(string id, string softwareName, DateTime eventDate, string eventType)
        {
            Id = id;
            SofwareName = softwareName;
            EventDate = eventDate;
            EventType = eventType;
        }
        #endregion

        #region Methods
        /// <summary>
        /// Возвращает строку, содержащую информацию о реестровой записи
        /// </summary>
        /// <returns>
        /// <para>Строка формата:</para>
        /// <para>Номер реестровой записи</para>
        /// <para>Наименование ПО</para>
        /// <para>Дата регистрации</para>
        /// <para>Тип События</para>
        /// </returns>
        public override string ToString()
        {
            return $"Номер реестровой записи: {Id}\n" +
                   $"Наименование ПО: {SofwareName}\n" +
                   $"Дата регистрации: {GetEventDateToString("dd.MM.yy")}\n" +
                   $"Тип события: {EventType}";
        }
        /// <summary>
        /// Устанавливает дату в EventDate из входящей строки
        /// </summary>
        /// <param name="date">Строчная дата формата дд.ММ.гг</param>
        public void SetEventDateByString(string date)
        {
            int[] splittedDate = date.Split('.').Select(x => int.Parse(x)).ToArray();
            EventDate = new DateTime(splittedDate[2], splittedDate[1], splittedDate[0]);
        }
        /// <summary>
        /// Возвращает строку, содержащую дату из EventDate
        /// </summary>
        /// <param name="format">Необходимый формат даты</param>
        /// <returns>Строка, содержащая дату в необходимом формате</returns>
        public string GetEventDateToString(string format)
        {
            return EventDate.ToString(format);
        }
        #endregion
    }
    /// <summary>
    /// Класс, содержащий информацию об отчёте: период, за который сформирован отчёт, а также все записи за данный период
    /// </summary>
    public class Report
    {
        public DateTime BeginningDate { get; set; }
        public DateTime EndingDate { get; set; }
        public List<RegisrtyWrite> RegisrtyWrites { get; set; }
        /// <summary>
        /// Создание объекта класса Report
        /// </summary>
        /// <param name="beginningDate">Начало периода, с которого берутся записи</param>
        public Report(DateTime beginningDate)
        {
            BeginningDate = beginningDate;
            EndingDate = DateTime.Today;
            RegisrtyWrites = new List<RegisrtyWrite>();
        }
        /// <summary>
        /// Возвращает оформленный отчёт с указанием периода и информацией о каждой записи за этот период
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine($"Отчёт за период с {BeginningDate.ToString("dd.MM.yy")} по {EndingDate.ToString("dd.MM.yy")}");
            for (int i = 0; i < RegisrtyWrites.Count; i++)
            {
                stringBuilder.AppendLine($"({i+1})");
                stringBuilder.AppendLine(RegisrtyWrites[i].ToString());
                stringBuilder.AppendLine();
            }
            return stringBuilder.ToString();
        }
        /// <summary>
        /// Добавление новой записи в список записей
        /// </summary>
        /// <param name="regisrtyWrite">Новая запись</param>
        public void Add(RegisrtyWrite regisrtyWrite) => RegisrtyWrites.Add(regisrtyWrite);
    }
    #endregion

    #region SOURCE
    /// <summary>
    /// Интерфейс для реализации получения информации из источника
    /// </summary>
    public interface IGetContent
    {
        /// <summary>
        /// Получение информации из источника
        /// </summary>
        /// <param name="source">Источник</param>
        XLWorkbook GetContent(string source);
    }

    /// <summary>
    /// Класс получения XLSX по ссылке
    /// </summary>
    public class GetDataFromUrl : IGetContent
    {
        /// <summary>
        /// Получает XLSX по ссылке
        /// </summary>
        /// <param name="source">Источник, из готорого берётся XLSX</param>
        /// <returns>Полученный файл XLSX</returns>
        public XLWorkbook GetContent(string source)
        {
            // ссылка на скачивание xlsx
            var url = source + (source.IndexOf("?") >= 0 ? "&" : "?") + "export=list";
            // место назначения файла
            var destinationPath = Path.Combine(Environment.CurrentDirectory, "..", "..", "..", "ExportedRegister.xlsx");

            // скачиваем файл
            DownloadXlsx(url, destinationPath).GetAwaiter().GetResult();

            return new XLWorkbook(destinationPath);
        }
        /// <summary>
        /// Асинхронное скачивание XLSX файла из источника
        /// </summary>
        /// <param name="url">Ссылка-источник</param>
        /// <param name="destinationPath">Путь назначения файла</param>
        /// <returns></returns>
        private async Task DownloadXlsx(string url, string destinationPath)
        {
            using(var client = new HttpClient())
            {
                // Скачиваем файл
                var response = await client.GetAsync(url);
                response.EnsureSuccessStatusCode();

                // Сохраняем файл на диск
                using (var fileStream = new FileStream(destinationPath, FileMode.Create))
                {
                    await response.Content.CopyToAsync(fileStream);
                }
            }
        }
    }
    /// <summary>
    /// Класс получения данных из файла
    /// </summary>
    public class GetDataFromFile : IGetContent
    {
        /// <summary>
        /// Получает XLSX по пути к файлу
        /// </summary>
        /// <param name="source">Источник, из готорого берётся XLSX</param>
        /// <returns>Полученный файл XLSX</returns>
        public XLWorkbook GetContent(string source)
        {
            return new XLWorkbook(source);
        }
    }
    #endregion

    #region PARSER
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
                if(IsSheetEnded(currentLine))
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
    #endregion

    #region STORE
    /// <summary>
    /// Класс сохранения отчёта в файл
    /// </summary>
    public class Store
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
    #endregion

    #region PROCESS
    /// <summary>
    /// Класс, управляющий логикой приложения
    /// </summary>
    public class Process
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
                if(IsValidUrl(source))
                {
                    var dataCollecor = new GetDataFromUrl();
                    _workbook = dataCollecor.GetContent(source);
                }
                else if(IsValidInputPath(source))
                {
                    var dataCollecor = new GetDataFromFile();
                    _workbook = dataCollecor.GetContent(source);
                }
                // если по источнику получили файл xlsx
                if (_workbook != null)
                {
                    // ищем подходящие записи и сохраняем в _report
                    var converter = new Converter(_workbook.Worksheet("Данные"));
                    _report = converter.GenerateReport(beginningDate);

                    // сохраняем полученный отчёт
                    var store = new Store();
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
            return path.IndexOfAny(Path.GetInvalidPathChars()) == -1;
        }
    }
    #endregion
    internal class Program
    {
        static void Main(string[] args)
        {
            int[] splittedDate = args[0].Split('.').Select(x => int.Parse(x)).ToArray();
            DateTime beginningDate = new DateTime(splittedDate[2], splittedDate[1], splittedDate[0]);

            var process = new Process();
            // дата начала, источник, выходной файл
            process.Start(beginningDate, args[1], args[2]);

            //var process = new Process();
            //process.Start(new DateTime(2024, 12, 28),
            //              "https://reestr.digital.gov.ru/reestr/",
            //              "C:\\Users\\shtoretc\\Downloads\\result.txt");
        }
    }
}
