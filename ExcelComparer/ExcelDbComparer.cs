using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Text.RegularExpressions;

namespace ExcelComparer
{
    class ExcelDbComparer
    {
        #region FIELDS

        private string file;
        private string[] fileList;
        private int FOIColumnIndex = 0;
        private string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;extended properties =\"excel 8.0;hdr=no;IMEX=1\";data source={0}";
        private string sheenName = "Лист1$";
        private List<NameFromDataTable> allDublicates = new List<NameFromDataTable>();
        private DataTable dataTableForAnalyze;

        #endregion

        public ExcelDbComparer(string file, string[] list)
        {
            if (String.IsNullOrEmpty(file) || list == null || list.Length == 0)
            {
                throw new ArgumentException("Неверные аргументы");
            }
            this.file = file;
            this.fileList = list;
        }

        public void RunWork()
        {
            LogWriter.Write("Анализируемый файл: " + this.file);
            dataTableForAnalyze = ReadExcelFile(this.file);
            List<NameFromDataTable> namesForAnalyze = ReadNamesFromData(dataTableForAnalyze);

            foreach (string f in this.fileList)
            {
                LogWriter.Write("Поиск дубликатов в файле: " + f);
                DataTable dtBase = ReadExcelFile(f);
                List<NameFromDataTable> namesBase = ReadNamesFromData(dtBase);
                List<NameFromDataTable> dublicates = GetDuplicateList(namesForAnalyze, namesBase);
                this.allDublicates.AddRange(dublicates);
            }

            LogWriter.Write("Всего найдено: " + this.allDublicates.Count);

            int before = this.dataTableForAnalyze.Rows.Count;
            DeleteDublicates();
            int after = this.dataTableForAnalyze.Rows.Count;
            LogWriter.Write("Удалено: " + (before - after));
            LogWriter.Write("Уникальных записей:" + this.dataTableForAnalyze.Rows.Count);

            this.dataTableForAnalyze.WriteXml("my.xlsx");
        }

        private void DeleteDublicates()
        {
            foreach (NameFromDataTable dubl in this.allDublicates)
            {
                this.dataTableForAnalyze.Rows[dubl.ID].Delete();
            }
            this.dataTableForAnalyze.AcceptChanges();
        }

        /// <summary>
        /// Возвращает список дубликатов
        /// </summary>
        /// <param name="list1">Анализируемый список</param>
        /// <param name="list2">Список для сравнения</param>
        /// <returns>Список дубликатов из первого списка, которые есть во втором</returns>
        public List<NameFromDataTable> GetDuplicateList(List<NameFromDataTable> list1, List<NameFromDataTable> list2)
        {
            List<NameFromDataTable> result = new List<NameFromDataTable>();
            foreach (NameFromDataTable name1 in list1)
            {
                foreach (NameFromDataTable name2 in list2)
                {
                    if(name1.Name == name2.Name)
                    {
                        result.Add(name1);
                        string message = String.Format("Найден дуликат: {0} : {1} = {2} : {3}", name1.ID, name2.Name, name2.ID, name2.Name);
                        LogWriter.Write(message);
                    }
                }
            }
            LogWriter.Write("Всего найдено дубликатов в файле: " + result.Count);
            return result;
        }

        /// <summary>
        /// Прочитать файл Excel в DataTable
        /// </summary>
        /// <param name="file">Путь к фалу</param>
        /// <returns>DataTable с содержимым файла</returns>
        public DataTable ReadExcelFile(string file)
        {
            try
            {
                /*
                var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", this.file);
                */
                string connectionString = String.Format(this.connectionString, file);
                string selectQuery = String.Format("SELECT * FROM [{0}]", this.sheenName);
                OleDbDataAdapter adapter = new OleDbDataAdapter(selectQuery, connectionString);
                DataSet ds = new DataSet();
                adapter.Fill(ds, "Base");
                DataTable data = ds.Tables["Base"];

                LogWriter.Write(String.Format("Прочитан файл {0}. Столбцов: {1}. Строк: {2}", file, data.Columns.Count, data.Rows.Count));

                return data;
            }
            catch (Exception e)
            {
                LogWriter.Write("Не удалось прочитать файл: " + e.Message);
                throw new ComparerException("Не удалось прочитать файл: " + file);
            }
        }

        /// <summary>
        /// Чтение списка ФИО и ИД записи
        /// Столбец с ФИО по умолчанию FOIColumnIndex = 0
        /// </summary>
        /// <param name="data">DataTable с данными</param>
        /// <returns>Список NameFromDataTable (Номер записи - ФИО)</returns>
        public List<NameFromDataTable> ReadNamesFromData(DataTable data)
        {
            List<NameFromDataTable> result = new List<NameFromDataTable>();
            int id = 0;
            foreach (DataRow row in data.Rows)
            {
                string rawName = row[this.FOIColumnIndex].ToString();
                string fio = ProcessName(rawName.Trim().ToLower());
                if (fio != null)
                {
                    result.Add(new NameFromDataTable() { ID = id, Name = fio });
                }
                else
                {
                    LogWriter.Write(String.Format("Запись: {0}. ФИО: {1} - Не валидна!", id, 
                        !String.IsNullOrEmpty(row[this.FOIColumnIndex].ToString()) ? row[this.FOIColumnIndex].ToString() : "<Нет значения>" ));
                }
                id++;
            }

            LogWriter.Write("Прочитано валидных имён из файла: " + result.Count);

            return result;
        }

        /// <summary>
        /// Обработка ФИО. Фильтрация пустых, невалидный записей, удаление дубликатов в ФИО
        /// </summary>
        /// <param name="name">ФИО</param>
        /// <returns>Обработанное ФИО</returns>
        private string ProcessName(string name)
        {
            string result = null;
            if(!String.IsNullOrEmpty(name))
            { 
                name = Regex.Replace(name, @"\W+", " ");                      //удаление мусора (не цифрово-алфовитные символы)
                name = Regex.Replace(name, @"\s+", " ");                      //замена двойных пробелов
                string[] words = name.Split(' ');                             //разбиваем на слова
                if(words.Length >= 2)
                { 
                    if (words.Length > 3)                                     //удалим дубликаты в фио (опасно to-do)
                    {
                        words = words.Distinct<string>().ToArray();
                    }
                    result = String.Join(" ", words);
                }
            }
            return result;
        }
    }

    struct NameFromDataTable
    {
        public int ID { get; set; }
        public string Name { get; set; }
    }
}