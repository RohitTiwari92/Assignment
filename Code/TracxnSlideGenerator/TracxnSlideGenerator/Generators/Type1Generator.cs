using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using TracxnSlideGenerator.Model;

namespace TracxnSlideGenerator.Generators
{
   public class Type1Generator
    {

        Application ppApp = null;
        string templatepath = "";
        string temppath = "";

        public string BuildType1SLides(string Path, string Data,bool Update =false)
        {
            string Result = "";
            if(File.Exists(Path))
            {
                return "File Already exists";
            }
            string Spath = SecondFilePath(Path);
            if (File.Exists(Spath))
            {
                Spath = SecondFilePath(Spath);
            }
            string Crs = CreateFile(Path);
            if (!Crs.Equals("success"))
            {
                return Crs;
            }
            Type1Model Model = new Type1Model();
            try
            {
              Model=  JsonConvert.DeserializeObject<Type1Model>(Data);
            }
            catch
            {
                return "Please Check  your Input (Json Data) ";
            }
            UpdateSlide(Path, Model,Update);
          
           
  
       
            return Result;
        }

        string SecondFilePath(string Path)
        {
            string SPath = "";
            FileInfo info = new FileInfo(Path);
            string time = DateTime.Now.ToString("ddMMyyyyTHHmmss");
            string name = System.IO.Path.GetFileNameWithoutExtension(Path);
            name = name + "_" + time;
            SPath = info.DirectoryName + @"\" + name + info.Extension;
            return SPath;
        }
       
        string CreateFile(string Path)
        {
            if(File.Exists(Path))
            {
                return "File Already exists";
            }

            using (Stream output = File.Create(Path))
            {
               
               Stream input = new MemoryStream(TracxnSlideGenerator.Properties.Resources.Type1);
               input.CopyTo(output);
               input.Close();
                //sw.WriteLine(TracxnSlideGenerator.Properties.Resources.Type1);
            }
             temppath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "PPT_Rohit_1.pptx");

             using (Stream output = File.Create(temppath))
             {

                 Stream input = new MemoryStream(TracxnSlideGenerator.Properties.Resources.Type1);
                 input.CopyTo(output);
                 input.Close();
                 //sw.WriteLine(TracxnSlideGenerator.Properties.Resources.Type1);
             }

            templatepath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "Chart1.crtx");

            using (Stream output = File.Create(templatepath))
            {

                Stream input = new MemoryStream(TracxnSlideGenerator.Properties.Resources.Chart1);
                input.CopyTo(output);
                input.Close();
                //sw.WriteLine(TracxnSlideGenerator.Properties.Resources.Type1);
            }

        
            return "success";
        }

        void UpdateSlide(string path , Type1Model Model,bool update )
        {
            _Presentation PPTObject = getPowerpointObject(path, update);
            if(PPTObject!=null)
            {
                UpdateTable(path, Model, PPTObject);
                UpdateTitle(PPTObject, Model,Model.Table_Color);
            }
            try
            {
                PPTObject.Save();
                PPTObject.Close();
                ppApp.Quit();
                NAR(PPTObject);
                NAR(ppApp);
                PPTObject = null;
                ppApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers(); 
                FreeFIle(path);
            }
            catch(Exception ex)
            {
                //log here 
            }
        }

        private void NAR(object o)
        {
            try
            {
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0) ;
            }
            catch { }
            finally
            {
                o = null;
            }
        }
        void FreeFIle(string Path)
        {

            Process tool = new Process();
            tool.StartInfo.FileName = "handle.exe";
            tool.StartInfo.Arguments = Path + " /accepteula";
            tool.StartInfo.UseShellExecute = false;
            tool.StartInfo.RedirectStandardOutput = true;
            tool.Start();
            tool.WaitForExit();
            string outputTool = tool.StandardOutput.ReadToEnd();

            string matchPattern = @"(?<=\s+pid:\s+)\b(\d+)\b(?=\s+)";
            foreach (Match match in Regex.Matches(outputTool, matchPattern))
            {
                Process.GetProcessById(int.Parse(match.Value)).Kill();
            }
        }
       void UpdateTitle(_Presentation PPTObject, Type1Model Model,int color)
        {
            foreach(Slide slide in PPTObject.Slides)
            {
              Microsoft.Office.Interop.PowerPoint.Shape shape=  GetShape(slide, "Title");
              string text = shape.TextFrame.TextRange.Text;
              text = text.Remove(29);
                 text = text +" "+ Model.AnalystName;
              //text = text + " " + "rohit";
              shape.TextFrame.TextRange.Text = text;
              shape.TextFrame.TextRange.Font.Color.RGB = color;
            }
        }

        void AddSlide(_Presentation PPTObject,int i )
        {
            PPTObject.Slides.InsertFromFile(temppath,i );
          //  PPTObject.Save();
        }

         void UpdateTable(string path, Type1Model Model, _Presentation PPTObject)
        {
             Slides SL=  PPTObject.Slides;
             int i = 1;
                 Slide sl0 = SL[i];
                 Microsoft.Office.Interop.PowerPoint.Shape shape = GetShape(SL[i], "Table");
                 Microsoft.Office.Interop.PowerPoint.Table tbl = shape.Table;
                 List<chartclass> ChartData = new List<chartclass>();
                 UpdateTableColor(Model, tbl);
                 int y = 2;
                 for (int index = 0; index < Model.DataList.Count; index++)
                 {
                     Type1SubModel rowdat = Model.DataList[index];
                     Row row = tbl.Rows.Add();
                     y++;
                     FillDataInROw(rowdat, row);
                     for (int temp = 1; temp < 7; temp++)
                     {
                         tbl.Cell(y, temp).Shape.Fill.ForeColor.RGB = 16777215;
                     }
                     tbl.Background.Fill.BackColor.RGB = 0;
                     chartclass cobj = new chartclass();
                     cobj.count = rowdat.Number_of_Companies;
                     cobj.country = rowdat.Markets;
                     float per = (shape.Height / PPTObject.PageSetup.SlideHeight) * 100;
                     ChartData.Add(cobj);
                     if (index == Model.DataList.Count-1)
                     {
                         
                         while (true)
                         {
                             float per_Temp = (shape.Height / PPTObject.PageSetup.SlideHeight) * 100;
                             if (per_Temp > 60)
                             {
                                 break;
                             }
                             row = tbl.Rows.Add();
                             chartclass cobj_Temp = new chartclass();
                             cobj_Temp.count = 0;
                             cobj_Temp.country = "";

                             ChartData.Add(cobj_Temp);
                         }
                   
                     }
                     if (per > 60)
                     {
                         y = 2;
                         UpdateChart(sl0, ChartData,Model.Table_Color);
                         ChartData = new List<chartclass>();
                         AddSlide(PPTObject, i);
                         i++;
                         sl0 = SL[i];
                         shape = GetShape(SL[i], "Table");
                         tbl = shape.Table;
                         UpdateTableColor(Model, tbl);
                     }
                 }
             UpdateChart(sl0, ChartData,Model.Table_Color);
             
        }

         private static void UpdateTableColor(Type1Model Model, Microsoft.Office.Interop.PowerPoint.Table tbl)
         {
             for (int temp = 1; temp < 7; temp++)
             {
                 tbl.Cell(2, temp).Shape.Fill.ForeColor.RGB = Model.Table_Color;
             }
         }

        void UpdateChart(Slide sl0, List<chartclass> ChartData,int color)
        {
             ChartData.Reverse();
            object misValue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.PowerPoint.Shape chartshape = GetShape(sl0, "Chart");
            chartshape.Chart.ApplyDataLabels(Microsoft.Office.Interop.PowerPoint.XlDataLabelsType.xlDataLabelsShowValue);
            
            dynamic workbook =  chartshape.Chart.ChartData.Workbook;
            dynamic worksheet = workbook.Worksheets(1);
            int x = 2;
            foreach (chartclass cc in ChartData)
            {
                worksheet.Cells[x, 1] = cc.country;
                worksheet.Cells[x, 2] = cc.count;
                x++;
            }


            chartshape.Chart.SeriesCollection(1).DataLabels(1).Text = ChartData [0].count;

             int i = 1;
            foreach (chartclass cc in ChartData)
            {
                chartshape.Chart.SeriesCollection(1).Points(i).Format.Fill.ForeColor.RGB = color; 
  
                i++;
            }
            
           // workbook.Save();
           // workbook.Close(true, misValue, misValue);
            workbook.Application.Quit();
        }
        private static void FillDataInROw(Type1SubModel rowdat, Row row)
        {
            
            int k = 1;
            Microsoft.Office.Interop.PowerPoint.Cell cell_ = row.Cells[k];
            Microsoft.Office.Interop.PowerPoint.TextFrame tf = cell_.Shape.TextFrame;
            tf.TextRange.Text = rowdat.Markets;
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Name = rowdat.Markets_CFont.FontName;
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Color.RGB = rowdat.Markets_CFont.Color;
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Size = rowdat.Markets_CFont.FontSize;

            cell_ = row.Cells[k + 2];
            tf = cell_.Shape.TextFrame;
            tf.TextRange.Text = rowdat.Companies.Company + "\n";
            int length = tf.TextRange.Text.Length;
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Name = rowdat.Companies.Company_CFont.FontName;
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Color.RGB = rowdat.Companies.Company_CFont.Color;
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Size = rowdat.Companies.Company_CFont.FontSize;
            tf.TextRange.Text = tf.TextRange.Text + rowdat.Companies.State;
            int textlength = rowdat.Companies.State.Count();
            tf.TextRange.Paragraphs(length, textlength).Font.Name = rowdat.Companies.State_CFont.FontName;
            tf.TextRange.Paragraphs(length, textlength).Font.Color.RGB = rowdat.Companies.State_CFont.Color;
            tf.TextRange.Paragraphs(length, textlength).Font.Size = rowdat.Companies.State_CFont.FontSize;

            cell_ = row.Cells[k + 3];
            tf = cell_.Shape.TextFrame;
            tf.TextRange.Text = rowdat.crawled.Tracked.ToString();
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Name = rowdat.crawled.Tracked_CFont.FontName;
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Color.RGB = rowdat.crawled.Tracked_CFont.Color;
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Size = rowdat.crawled.Tracked_CFont.FontSize;

            cell_ = row.Cells[k + 4];
            tf = cell_.Shape.TextFrame;
            tf.TextRange.Text = rowdat.crawled.Missed.ToString();
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Name = rowdat.crawled.Missed_CFont.FontName;
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Color.RGB = rowdat.crawled.Missed_CFont.Color;
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Size = rowdat.crawled.Missed_CFont.FontSize;

            cell_ = row.Cells[k + 5];
            tf = cell_.Shape.TextFrame;
            tf.TextRange.Text = rowdat.crawled.Total_Companies.ToString();
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Name = rowdat.crawled.Total_Companies_CFont.FontName;
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Color.RGB = rowdat.crawled.Total_Companies_CFont.Color;
            tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Size = rowdat.crawled.Total_Companies_CFont.FontSize;
        }

      public  Microsoft.Office.Interop.PowerPoint.Shape GetShape(Slide slide,string reqobj)
        {
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
            {
                if (shape.Name.Contains(reqobj))
                {
                    return shape;
                }
            }
            return null;
        }

        _Presentation getPowerpointObject(string path, bool update)
        {

            ppApp = new Application();
            ppApp.Visible = MsoTriState.msoTrue;
            Presentations oPresSet = ppApp.Presentations;
            if (!update)
            {
                _Presentation PPTObject = oPresSet.Open(path, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
                return PPTObject;
            }
            else
            {
                _Presentation PPTObject = oPresSet.Open(path, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                return PPTObject;
            }
        }


    }
}
