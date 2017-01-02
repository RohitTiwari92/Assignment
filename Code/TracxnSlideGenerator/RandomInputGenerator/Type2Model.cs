using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RandomInputGenerator
{
    class MasterModel
    {
        public int Title_Color { get; set; }
        public List<Type2Model> Data { get; set; }
    }

    class Type2Model
    {
        public Type2SUbModel Header { get; set; }

        public int header_Color { get; set; }
        public List<Type2SUbModel> PlayerName { get; set; }
    }

    class Type2SUbModel
    {
        public string Name { get; set; }
        public CFont Name_CFont { get; set; }
    }
}
