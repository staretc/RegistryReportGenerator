using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace Source
{
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
            var destinationPath = Path.Combine(Environment.CurrentDirectory, "..", "ExportedRegister.xlsx");

            // скачиваем файл
            DownloadXlsx(url, destinationPath).GetAwaiter().GetResult();

            // отправляем уведомление об успешном скачивании файла
            Console.WriteLine("Файл xlsx успешно получен!");

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
            using (var client = new HttpClient())
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
}
