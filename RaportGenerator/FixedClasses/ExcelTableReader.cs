using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RaportGenerator.FixedClasses
{
    public static class ExcelTableReader
    {
        // МЕТОД ЗАГРУЗКИ ДАННЫХ ИЗ EXCEL ТАБЛИЦЫ
        public static DataTableCollection Load(string path)
        {
            // Открываем поток для считывания файла
            using (FileStream fs = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                // Открываем поток для считывания Excel-файла в потоке считывания файла
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(fs))
                {
                    // Загружаем все считанные данные в датасет
                    DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (x) => new ExcelDataTableConfiguration() { }
                    });

                    return db.Tables;
                }
            }
        }

    }
}
