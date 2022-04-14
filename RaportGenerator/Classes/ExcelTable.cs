using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RaportGenerator.Classes
{
    /// <summary>
    /// Класс, представляющий Excel таблицу, расположенную на конкретном листе.
    /// </summary>
    public class ExcelTable
    {
        // Название таблицы (причина смертности)
        public string? TableName {get;set;}

        // Список данных (города со статистикой)
        public List<Entry> Entries = new List<Entry>();

        // Период
        public string? DatesLast { get; set; }
        public string? DatesCurrent { get; set; }

    }
}
