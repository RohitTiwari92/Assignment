using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RandomInputGenerator
{
   public class InputGenerator
    {
       public string GenrateDataForPPtType2(int Rows)
        {
            Type2ModelJsonDataGenerator pobj = new Type2ModelJsonDataGenerator();
            MasterModel data = pobj.GetModelDataForType2(Rows);
            string jres = pobj.CreateJsonForType1(data);
            return jres;
        }
       public string GenrateDataForPPtType1(int Rows)
       {
           Type1ModelJsonDataGenerator pobj = new Type1ModelJsonDataGenerator();
           Type1Model data = pobj.GetModelDataForType1(Rows);
           string jres = pobj.CreateJsonForType1(data);
           return jres;
       }
    }
}
