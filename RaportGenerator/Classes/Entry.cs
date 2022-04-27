using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RaportGenerator.Classes
{
    /// <summary>
    /// Класс, представляющий основные данные таблицы (города с их статистикой).
    /// </summary>
    public class Entry
    {
        // Наименование объекта
        public string? Name { get; set; }

        // Свойства объектов
        public List<IProperty> Properties = new List<IProperty>();



        //// Наименование данных
        //public string? Name { get; set; }

        //// Количественные данные за текущий период 
        //public double QuantityCurrent { get; set; }

        //// Количественные данные за прощедший период
        //public double QuantityLast { get; set; }

        //// Числовая сравнительная динамика
        //public double? DynamicAbs { get; set; }

        //// Процентная сравнительная динамика
        //public double? DynamicPersents { get; set; }


    }
}
