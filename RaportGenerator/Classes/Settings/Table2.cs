using System.Collections.Generic;

namespace RaportGenerator.Classes.Settings
{
    public class Table2 : ITable
    {
        public List<int> TableName { get; set; }
        public List<int> DataName { get; set; }
        public Properties Properties { get; set; }
    }
}
