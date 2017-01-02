using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using TracxnSlideGenerator.Properties;
using TracxnSlideGenerator.Generators;

namespace TracxnSlideGenerator
{
    public class Generator
    {
        //get supported type
        public List<string> GetTypeOfMasterSlides()
        {
            List<string> Files = new List<string>();
            ResourceSet resourceSet = Properties.Resources.ResourceManager.GetResourceSet(CultureInfo.CurrentUICulture, true, true);
            foreach (DictionaryEntry entry in resourceSet)
            {
                object resourceName = entry.Key; 
                if ( resourceName.ToString().Contains( "Type"))
                {
                    Files.Add(resourceName.ToString());
                }
            }
            return Files;
        }

        public string BuildSlide(string type ,string path, string Data,bool update =false)
        {
            string result = "";
            try
            {
                if (type.Equals("Type1"))
                {
                    Type1Generator gen = new Type1Generator();
                    result = gen.BuildType1SLides(path, Data,update);
                }
                if (type.Equals("Type2"))
                {
                    Type2Generator gen = new Type2Generator();
                    result = gen.BuildType2SLides(path, Data,update);
                }
            }
            catch(Exception ex)
            {
                result = ex.Message;
            }
            return result;
        }

    }
}
