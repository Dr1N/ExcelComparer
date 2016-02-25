using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System.IO;

namespace ExcelComparer
{
    class ExcelDbComparer
    {
        #region FIELDS

        public FIELD_TYPE FieldType = FIELD_TYPE.NONE;

        private string file;
        private string[] fileList;
        private int FOIColumnIndex = 0;
        private string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;extended properties =\"excel 8.0;hdr=no;IMEX=1\";data source={0}";
        private string sheetName = "Лист1";
        private string currentFile = "";

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

        public void RunWork(object frm)
        {
            MainForm form = frm as MainForm;
            if(form == null)
            {
                LogWriter.Write("Невозможно начать обработку. Нет ссылки на главную форму");
                return;
            }
            if (FieldType == FIELD_TYPE.NONE)
            {
                LogWriter.Write("Не указан тип поля");
                return;
            }
            try
            {
                form.StartWork();

                LogWriter.Write("Анализируемый файл: " + this.file);
                LogWriter.Write("Файлы для сравнения:");
                foreach (string f in this.fileList)
                {
                    LogWriter.Write(f);
                }
                
                DataTable dtForAnalyze = ReadExcelFile(this.file);
                List<FieldData> fieldsForAnalyze = ReadFieldFromData(dtForAnalyze);
                List<FieldData> allDublicates = new List<FieldData>();
                foreach (string f in this.fileList)
                {
                    this.currentFile = f;
                    LogWriter.Write("Поиск дубликатов в файле: " + f);
                    DataTable dtBase = ReadExcelFile(f);
                    List<FieldData> fieldsForCompare = ReadFieldFromData(dtBase);
                    List<FieldData> dublicates = GetDuplicateList(fieldsForAnalyze, fieldsForCompare);
                    allDublicates.AddRange(dublicates);
                }

                LogWriter.Write("Всего найдено дубликатов: " + allDublicates.Count);

                DeleteDublicates(dtForAnalyze, allDublicates);

                LogWriter.Write("Дубликаты удалены");

                SaveDataTableInExcelFormat(dtForAnalyze);

                LogWriter.Write("Результат сохранён: " + Path.GetFullPath(Path.GetFileNameWithoutExtension(this.file) + "_filtered" + Path.GetExtension(this.file)));
            }
            catch(Exception e)
            {
                LogWriter.Write("Error! Не удалось обработать файл: " + e.Message);
            }
            finally
            {
                form.EndWork();
            }
        }

        /// <summary>
        /// Удалить дубликаты из DataTable
        /// </summary>
        /// <param name="dt">Обрабатываемый DataTable</param>
        /// <param name="dublicates">Список добликатов</param>
        private void DeleteDublicates(DataTable dt, List<FieldData> dublicates)
        {
            foreach (FieldData dubl in dublicates)
            {
                dt.Rows[dubl.ID].Delete();
            }
            dt.AcceptChanges();
        }

        /// <summary>
        /// Сохранить DataTable в формате Excel
        /// </summary>
        /// <param name="dt">Сохраняемый DataTable</param>
        private void SaveDataTableInExcelFormat(DataTable dt)
        {
            XLWorkbook woorkBook = new XLWorkbook();
            woorkBook.Worksheets.Add(dt, this.sheetName);
            string name = Path.GetFileNameWithoutExtension(this.file) + "_filtered" + Path.GetExtension(this.file);
            if(File.Exists(name))
            {
                File.Delete(name);
            }
            woorkBook.SaveAs(name);
        }

        /// <summary>
        /// Возвращает список дубликатов
        /// </summary>
        /// <param name="analyzed">Анализируемый список</param>
        /// <param name="forcompare">Список для сравнения</param>
        /// <returns>Список дубликатов из первого списка, которые есть во втором</returns>
        private List<FieldData> GetDuplicateList(List<FieldData> analyzed, List<FieldData> forcompare)
        {
            List<FieldData> result = new List<FieldData>();
            foreach (FieldData analizedField in analyzed)
            {
                foreach (FieldData fieldForCompare in forcompare)
                {
                    if(analizedField.Equals(fieldForCompare))
                    {
                        result.Add(analizedField);
                        string message = String.Format("Найден дуликат: {0} : {1} = {2} : {3}", analizedField.ID, fieldForCompare.FieldValue, fieldForCompare.ID, fieldForCompare.FieldValue);
                        LogWriter.Write(message);
                    }
                }
            }
            LogWriter.Write(String.Format("Всего найдено дубликатов в файле [{0}]: {1} ", this.currentFile, result.Count));
            return result;
        }

        /// <summary>
        /// Прочитать файл Excel в DataTable
        /// </summary>
        /// <param name="file">Путь к фалу</param>
        /// <returns>DataTable с содержимым файла</returns>
        private DataTable ReadExcelFile(string file)
        {
            LogWriter.Write(String.Format("Читаем файл: {0}", file));
            string connectionString = String.Format(this.connectionString, file);
            string selectQuery = String.Format("SELECT * FROM [{0}$]", this.sheetName);
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(selectQuery, connectionString))
            { 
                DataSet ds = new DataSet();
                adapter.Fill(ds, "Base");
                DataTable data = ds.Tables["Base"];

                LogWriter.Write(String.Format("Прочитан файл {0}. Столбцов: {1}. Строк: {2}", file, data.Columns.Count, data.Rows.Count));

                return data;
            }
        }

        /// <summary>
        /// Чтение списка ФИО и ИД записи
        /// Столбец с ФИО по умолчанию FOIColumnIndex = 0
        /// </summary>
        /// <param name="data">DataTable с данными</param>
        /// <returns>Список NameFromDataTable (Номер записи - ФИО)</returns>
        private List<FieldData> ReadFieldFromData(DataTable data)
        {
            LogWriter.Write("=======================================");
            LogWriter.Write("Читаем Поле из файла: " + this.currentFile);
            LogWriter.Write("=======================================");
            List<FieldData> result = new List<FieldData>();
            int id = 0;
            foreach (DataRow row in data.Rows)
            {
                string rawField = row[this.FOIColumnIndex].ToString().ToLower().Trim();
                FieldData fd = null;
                switch (FieldType)
                {
                    case FIELD_TYPE.NAME:
                        fd = new NameField(id, rawField);
                        break;
                    case FIELD_TYPE.PHONE:
                        fd = new PhoneField(id, rawField);
                        break;
                    default:
                        LogWriter.Write("Не известный тип поля: " + FieldType);
                        break;
                }
                if (fd.FieldValue != null)
                {
                    result.Add(fd);
                }
                else
                {
                    LogWriter.Write(String.Format("Запись: {0}. Поле: {1} - Не валидна!", id, 
                        !String.IsNullOrEmpty(row[this.FOIColumnIndex].ToString()) ? row[this.FOIColumnIndex].ToString() : "<Нет значения>" ));
                }
                id++;
            }

            LogWriter.Write("Прочитано валидных полей из файла: " + result.Count);

            return result;
        }
    }

    abstract class FieldData : IEquatable<FieldData>
    {
        public int ID { get; set; }
        public string FieldValue { get; set; }

        public FieldData(int id, string value)
        {
            ID = id;
            FieldValue = ProcessField(value);
        }
        public abstract string ProcessField(string field);
        public abstract bool Equals(FieldData other);
    }

    class NameField : FieldData
    {
        public NameField(int id, string value) : base(id, value) { }

        public override bool Equals(FieldData other)
        {
            if (other is NameField == false)
            {
                return false;
            }
            return FieldValue == other.FieldValue;
        }
              
        public override string ProcessField(string field)
        {
            string result = null;
            if (!String.IsNullOrEmpty(field))
            {
                field = Regex.Replace(field, @"\W+", " ");          //удаление мусора (не цифрово-алфовитные символы)
                field = Regex.Replace(field, @"\s+", " ");          //замена двойных пробелов
                string[] words = field.Split(' ');                  //разбиваем на слова
                if (words.Length >= 2)
                {
                    if (words.Length > 3)                           //удалим дубликаты в фио (опасно to-do)
                    {
                        words = words.Distinct<string>().ToArray();
                    }
                    result = String.Join(" ", words);
                }
            }
            return result;
        }
    }

    class PhoneField : FieldData
    {
        public PhoneField(int id, string value) : base(id, value) { }

        public override bool Equals(FieldData other)
        {
            if (other is FieldData == false)
            {
                return false;
            }
            return FieldValue.EndsWith(other.FieldValue) || other.FieldValue.EndsWith(FieldValue);
        }

        public override string ProcessField(string field)
        {
            if (!String.IsNullOrEmpty(field))
            {
                field = Regex.Replace(field, @"\D+", "");              //удаление не цифр
                field = Regex.Replace(field, @"\s+", "");              //удаление пробелов
                if (field.Length > 6)
                {
                    return field;
                }
            }
            return null;
        }
    }

    enum FIELD_TYPE
    {
        NONE = 0,
        NAME = 1,
        PHONE = 2
    }
}