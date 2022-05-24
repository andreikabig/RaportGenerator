using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RaportGenerator.FixedClasses
{
    // КЛАСС, ПРЕДСТАВЛЯЮЩИЙ EXCEL ТАБЛИЦУ, РАСПОЛОЖЕННУЮ НА КОНКРЕТНОМ ЛИСТЕ
    public class ExcelTable
    {
        // НАЗВАНИЕ ТАБЛИЦЫ
        public string? Name { get; set; }

        // Список данных (города со статистикой)
        public List<Entry>? Entries { get; set; }

        // Период
        public string? DatesLast { get; set; }
        public string? DatesCurrent { get; set; }


        // МЕТОД ПОДСЧЕТА ИТОГО ПО СВОЙСТВУ ??????????????
        public List<Property>? GetStatistic()
        {
            // Статистика 
            List<Property>? stats = null;
            if (Entries != null)
            {
                stats = new List<Property>();

                // Выбираем доступные свойства (названия)
                foreach (var prop in Entries[0].Properties) 
                {
                    stats.Add(new Property() { Name = prop.Name });
                }

                // Считаем статистику для каждого найденного свойства
                var props = Entries.Select(x => x.Properties.Where(x => x.Name == "ff")).ToList();
            }

            return stats;
        }
    }
}
