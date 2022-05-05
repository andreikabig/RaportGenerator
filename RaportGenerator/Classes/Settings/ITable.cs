using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RaportGenerator.Classes.Settings
{
    public interface ITable
    {
        public List<int> TableName { get; set; }
        public List<int> DataName { get; set; }
        public Properties Properties { get; set; }
    }
}
