using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RandomInputGenerator
{
    class Type2ModelJsonDataGenerator
    {
        public  MasterModel GetModelDataForType2(int n)
        {
            MasterModel master = new MasterModel();
            List<Type2Model> Data = new List<Type2Model>();
            for (int i=0 ;i<n;i++)
            {
                Random r =new Random(i);
                Type2Model model = new Type2Model();
                model.Header = new Type2SUbModel();
                model.Header.Name = RandomString(r.Next(3, 9));
                Type1ModelJsonDataGenerator invocfont=new Type1ModelJsonDataGenerator();
                model.Header.Name_CFont = invocfont.GetCfontObject(i);
                r =new Random(i+2);
                r =new Random(i+4);
                model.header_Color = r.Next(0, 255);
                int getplayercount = r.Next(1, 12);
                model.PlayerName = new List<Type2SUbModel>();
                for( int j =0 ; j < getplayercount ; j++)
                {
                    Type2SUbModel submodel = new Type2SUbModel();
                    submodel.Name =  RandomString(r.Next(3, 12));
                    submodel.Name_CFont =  invocfont.GetCfontObject(i+7);
                    model.PlayerName.Add(submodel);
                }
                Data.Add(model);
            }
            master.Data = new List<Type2Model>();
            master.Data = Data;
            Random ran =new Random(10);
            master.Title_Color = ran.Next(100, 244);
            return master;
        }
        public string CreateJsonForType1(MasterModel Data)
        {
            string Jsondata = JsonConvert.SerializeObject(Data);
            return Jsondata;
        }
        public string RandomString(int length)
        {
            Random random = new Random();
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}
