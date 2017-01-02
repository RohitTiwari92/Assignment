using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TracxnSlideGenerator.Generators;
using TracxnSlideGenerator;

namespace TestBot
{
    class Getobjecttest
    {
        _Presentation PPTObject { get; set; }
        
        [SetUp]
        public void GetPPTObject()
        {
            string solution_dir = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.WorkDirectory));
            string path = @solution_dir + @"\TestFiles\ObjectTestFile.pptx";
            string absolute = Path.GetFullPath(path);
            Application ppApp = new Application();
            ppApp.Visible = MsoTriState.msoTrue;
            Presentations oPresSet = ppApp.Presentations;
            PPTObject = oPresSet.Open(@path, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
          
        }

        [TearDown]
        public void tearDown()
        {
            PPTObject.Close();
        }

        [Test]
        public void ObjectTest()
        {
            Type1Generator gen=new Type1Generator();
            Microsoft.Office.Interop.PowerPoint.Shape shape= gen.GetShape(PPTObject.Slides[1], "Table");
            Assert.AreEqual("Table 1", shape.Name);
        }
        [Test]
        public void fileExistsTest()
        {
            string solution_dir = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.WorkDirectory));
            string path = @solution_dir + @"\TestFiles\ObjectTestFile.pptx";
            Type1Generator gen = new Type1Generator();
            string output = gen.BuildType1SLides(path, "");
            Assert.AreEqual("File Already exists", output);
        }
         [Test]
        public void TypeCheck()
        {
            Generator gen = new Generator();
            List<string> types = gen.GetTypeOfMasterSlides();
            Assert.AreEqual(2, types.Count);
        }



    }
}
