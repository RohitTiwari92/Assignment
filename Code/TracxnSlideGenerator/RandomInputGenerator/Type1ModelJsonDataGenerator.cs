using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RandomInputGenerator
{
    class Type1ModelJsonDataGenerator
    {

        public string CreateJsonForType1(Type1Model Data)
        {
            string Jsondata = JsonConvert.SerializeObject(Data);
            return Jsondata;
        }

          public  Type1Model GetModelDataForType1(int n)
        {
            Type1Model Data = new Type1Model();
            Data.AnalystName = "Rohit";
            Data.Table_Color = 0;
            Data.DataList = new List<Type1SubModel>();

            for (int i = 0; i < n; i++)
            {
                Type1SubModel model = new Type1SubModel();
                model.Markets = "United State " + i;
                model.Markets_CFont = GetCfontObject(i);
                int Seed = (int)DateTime.Now.Ticks;
                Random r = new Random(Seed + i);
                model.Number_of_Companies = r.Next(1, 50);

                model.Companies = new Most_popular_Companies();
                model.Companies.Company = "Facebook " + i;
                model.Companies.Company_CFont = GetCfontObject(i);
                model.Companies.State = "Menlo Park " + i;
                model.Companies.State_CFont = GetCfontObject(i);

                model.crawled = new Total_Companies_crawled();
                Seed = (int)DateTime.Now.Ticks;
                r = new Random(Seed + i);
                model.crawled.Missed = r.Next(1, 10);
                model.crawled.Missed_CFont = GetCfontObject(i);
                model.crawled.Total_Companies = Math.Round(GetRandomNumber(0, 10), 1);
                model.crawled.Total_Companies_CFont = GetCfontObject(i);
                Seed = (int)DateTime.Now.Ticks;
                r = new Random(Seed + i);
                model.crawled.Tracked = r.Next(1, 500);
                model.crawled.Tracked_CFont = GetCfontObject(i);
                Data.DataList.Add(model);
            }
            return Data;
        }

        double GetRandomNumber(double minimum, double maximum)
        {
            int Seed = (int)DateTime.Now.Ticks;
            Random random = new Random(Seed);
            return random.NextDouble() * (maximum - minimum) + minimum;
        }


         public  CFont GetCfontObject(int i)
        {
            string[] fontname = { "Batang", "Cambria", "Arial" };
            int Seed = (int)DateTime.Now.Ticks;
            Random r = new Random(Seed + i);
            CFont font = new CFont();
            font.Color = r.Next(50, 200);
            font.FontSize = r.Next(5, 10);
            font.FontName = fontname[r.Next(0, 2)];
            return font;
        }
         public  string RandomString(int length)
         {
             Random random = new Random();
             const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
             return new string(Enumerable.Repeat(chars, length)
               .Select(s => s[random.Next(s.Length)]).ToArray());
         }
    }
}
