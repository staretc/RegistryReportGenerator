using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data
{
    /// <summary>
    /// Класс, хранящий необходимую информацию из записи реестра
    /// </summary>
    public class RegisrtyWrite : IComparable<RegisrtyWrite>
    {
        #region Fields
        private string _id;
        private string _softwareName;
        private DateTime _eventDate;
        private string _eventType;
        private string _applicationNumber;
        private string _stateRegistrationNumber;
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
        public string ApplicationNumber
        {
            get => _applicationNumber;
            set => _applicationNumber = value;
        }
        public string StateRegistrationNumber
        {
            get => _stateRegistrationNumber;
            set => _stateRegistrationNumber = value;
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
        /// <param name="applicationNumber">Номер заявления</param>
        /// <param name="stateRegistrationNumber">Номер гос. регистрации</param>
        public RegisrtyWrite(string id, string softwareName, DateTime eventDate, string eventType, string applicationNumber, string stateRegistrationNumber)
        {
            Id = id;
            SofwareName = softwareName;
            EventDate = eventDate;
            EventType = eventType;
            ApplicationNumber = applicationNumber;
            StateRegistrationNumber = stateRegistrationNumber;
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
        /// <para>Номер заявления</para>
        /// <para>Номер гос. регистрации (при наличии)</para>
        /// </returns>
        public override string ToString()
        {
            var stringBuilder = new StringBuilder();
            stringBuilder.Append($"Номер реестровой записи: {Id}\n" +
                                 $"Наименование ПО: {SofwareName}\n" +
                                 $"Дата регистрации: {GetEventDateToString("dd.MM.yy")}\n" +
                                 $"Тип события: {EventType}\n" +
                                 $"Номер заявления: {ApplicationNumber}");
            if (StateRegistrationNumber != "-1")
            {
                stringBuilder.Append($"\nНомер гос. Регистрации: {StateRegistrationNumber}");
            }
            return stringBuilder.ToString();
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
        /// <summary>
        /// Сравнение двух записей по датам записей (по Id записей в случае одинаковых дат)
        /// </summary>
        /// <param name="other">Запись, с которой ведётся сравнение</param>
        /// <returns>Результат сравнения</returns>
        public int CompareTo(RegisrtyWrite other)
        {
            if (EventDate.Equals(other.EventDate))
            {
                return other.Id.CompareTo(Id);
            }
            return other.EventDate.CompareTo(EventDate);
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
                stringBuilder.AppendLine($"({i + 1})");
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
        public void Sort() => RegisrtyWrites.Sort();
    }
}
