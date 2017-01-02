using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RandomInputGenerator
{
   public class Type1Model
    {
       public string AnalystName { get; set; }

       public int Table_Color { get; set; }
       public List<Type1SubModel> DataList { get; set; }

    }
  public  class Type1SubModel
    {
      public string Markets { get; set; }

      public CFont Markets_CFont { get; set; }
      public int Number_of_Companies { get; set; }

     // public int Number_of_Companies_RGB { get; set; }

      public Most_popular_Companies Companies { get; set; }
      public Total_Companies_crawled crawled { get; set; }
    }

    public class Most_popular_Companies
    {
        public string Company { get; set; }
        public CFont Company_CFont { get; set; }
        public string State { get; set; }
        public CFont State_CFont { get; set; }
    }

    public class Total_Companies_crawled
    {
        public int Tracked { get; set; }

        public CFont Tracked_CFont { get; set; }
        public int Missed { get; set; }
        public CFont Missed_CFont { get; set; }
        public double Total_Companies { get; set; }
        public CFont Total_Companies_CFont { get; set; }
    }

    public class CFont
    {
        public string FontName { get; set; }
        public int FontSize { get; set; }
        public int Color { get; set; }
    }

}
