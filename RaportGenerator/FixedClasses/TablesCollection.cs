using RaportGenerator.FixedClasses;
using RaportGenerator.Classes.Settings;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace RaportGenerator.FixedClasses
{
    // КЛАСС ПРЕДСТАВЛЯЮЩИЙ МОДЕЛЬ ВСЕХ ТАБЛИЦ
    public class TablesCollection
    {
        public TablesCollection(DataTableCollection dtc)
        {
            DTC = dtc;
        }

        public TablesCollection(DataTableCollection dtc, ITable settings)
        {
            DTC = dtc;
            Tables = GetTables(settings);
            SortTables();
        }

        // Считанная коллекция
        public readonly DataTableCollection DTC;

        // Список таблиц
        public List<ExcelTable>? Tables { get; private set; }

        // МЕТОД ПОЛУЧЕНИЯ ТАБЛИЦ ИЗ DTC
        private List<ExcelTable>? GetTables(ITable tableSettings) 
        {
            // Объявляем таблицы
            List<ExcelTable> tables = new List<ExcelTable>();


            //DataTable? table = tableCollection1[Convert.ToString(ComboBoxPages.SelectedItem)];
            foreach (DataTable table in DTC)
            {
                if (table != null)
                {
                    // Создание объекта таблицы
                    ExcelTable exTable = new ExcelTable();

                    // Добавление названия таблицы
                    exTable.Name = table.Rows[tableSettings.TableName[0]][tableSettings.TableName[1]].ToString();

                    // Формирование списка объектов таблицы
                    exTable.Entries = new List<Entry>();


                    // Заополнение списка объектов таблицы
                    for (int i = tableSettings.Properties.RangeData[0]; i <= tableSettings.Properties.RangeData[1]; i++) // РЕЗУЛЬТИРУЮЩУЮ БУДЕМ ВЫВОДИТЬ САМОСТОЯТЕЛЬНО, ПОСЛЕ СОРТИРОВКИ!
                    {
                        // Создание нового объекта
                        Entry entry = new Entry();

                        // Добавление названия объекта
                        entry.Name = table.Rows[i][tableSettings.DataName[1]].ToString();

                        /* СВОЙСТВА НЕ ЗАПОЛНЯЮТСЯ */
                        // Заполнение свойств объекта

                        int _j = tableSettings.Properties.RangePropertisNames.Range[0];// ПОЧЕМУ ПЕРЕДАЕТСЯ ПО ССЫЛКЕ ? 
                        int __j = tableSettings.Properties.RangePropertisNames.Range[1]; 
                        for (int j = _j; j <= __j; j++)
                        {


                            // Считывание названия свойства
                            string? propName = table.Rows[tableSettings.Properties.RangePropertisNames.Row][j].ToString();

                            // Считывание значения свойства
                            var value = DoubleConverter(table.Rows[i][j]);

                            // Объявление свойства
                            Property property;

                            // Если конвертер вернул double число
                            if (value != null)
                            {
                                // Создание свойства объекту с числом
                                property = new Property() { Name = propName, Value = (double)value };
                            }
                            else
                            {
                                // Создание свойства объекту без числа
                                property = new Property() { Name = propName };
                            }

                            // Если свойство есть
                            if (property != null)
                            {
                                // Добавляем свойство в список свойств
                                entry.Properties.Add(property);
                            }
                            else
                            {
                                throw new Exception(message: "Одно из свойств объекта не было сохранено корректно! Пожалуйста, проверьте правильность составленного документа!");
                            }


                        }

                        exTable.Entries.Add(entry);
                    }
                    tables.Add(exTable); // Создать модель пофикшеную
                }
            }

            // Если не null, то сохраняем
            if (tables != null)
                Tables = tables;

            // Возвращаем таблицы
            return tables;
        }
        private double? DoubleConverter(object obj)
        {
            try
            {
                return (double)obj;
            }
            catch (System.InvalidCastException ex)
            {
                return null;
            }
        }
        public void SortTables() {
            // СОРТИРОВКА ПО ПОСЛЕДНЕМУ СВОЙСТВУ
            if (Tables != null)
            {
                foreach (var table in Tables)
                    table.Entries = table.Entries.OrderBy(e => e.Properties[^1].Value).ToList();
            }
            
        }
    }
}
