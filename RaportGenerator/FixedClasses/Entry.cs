using System.Collections.Generic;

namespace RaportGenerator.FixedClasses
{
    // Класс, представляющий основные данные таблицы (города с их статистикой).
    public class Entry
    {
        // Наименование объекта
        public string? Name { get; set; }

        // Описание свойства
        public string? Description { get; set; }

        // Свойства объектов
        public List<Property> Properties = new List<Property>();

    }
}