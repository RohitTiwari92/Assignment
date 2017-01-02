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
    class Type2Generator
    {
        Application ppApp = null;
        string temppath = "";
        public string BuildType2SLides(string Path, string Data, bool Update = false)
        {
            string Result = "success";
            string Crs = CreateFile(Path);
            if (!Crs.Equals("success"))
            {
                return Crs;
            }
            MasterModel Model = new MasterModel();
            try
            {
                Model = JsonConvert.DeserializeObject<MasterModel>(Data);
            }
            catch
            {
                return "Please Check  your Input (Json Data) ";
            }
            UpdateSlide(Path, Model,Update);
            return Result;
        }
        void UpdateSlide(string path, MasterModel Model, bool Update )
        {
             
            _Presentation PPTObject = getPowerpointObject(path,Update);
            if (PPTObject != null)
            {
                UpdateTable(path, Model.Data, PPTObject);
                UpdateTitle(PPTObject, Model.Title_Color);
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
                // 
            }
            catch (Exception ex)
            {
                //log here 
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
        void UpdateTitle(_Presentation PPTObject,  int color)
        {
            foreach (Slide slide in PPTObject.Slides)
            {
                Microsoft.Office.Interop.PowerPoint.Shape shape = GetShape(slide, "Title");
                shape.TextFrame.TextRange.Font.Color.RGB = color;
            }
        }
        string CreateFile(string Path)
        {
            if (File.Exists(Path))
            {
                return "File Already exists";
            }

            using (Stream output = File.Create(Path))
            {

                Stream input = new MemoryStream(TracxnSlideGenerator.Properties.Resources.Type2);
                input.CopyTo(output);
                input.Close();
                //sw.WriteLine(TracxnSlideGenerator.Properties.Resources.Type1);
            }
            temppath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "PPT_Rohit_2.pptx");

            using (Stream output = File.Create(temppath))
            {

                Stream input = new MemoryStream(TracxnSlideGenerator.Properties.Resources.Type2);
                input.CopyTo(output);
                input.Close();
                //sw.WriteLine(TracxnSlideGenerator.Properties.Resources.Type1);
            }

            return "success";
        }
        _Presentation getPowerpointObject(string path, bool Update)
        {

            ppApp = new Application();
            ppApp.Visible = MsoTriState.msoTrue;
            Presentations oPresSet = ppApp.Presentations;
            if (!Update)
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
        void AddSlide(_Presentation PPTObject, int i)
        {
            PPTObject.Slides.InsertFromFile(temppath, i);
            //  PPTObject.Save();
        }

        void UpdateTable(string path, List<Type2Model> Model, _Presentation PPTObject)
        {
            Slides SL = PPTObject.Slides;
            int i = 1;
            Slide sl0 = SL[i];
            Microsoft.Office.Interop.PowerPoint.Shape shape = GetShape(SL[i], "Table");
            Microsoft.Office.Interop.PowerPoint.Table tbl = shape.Table;
            shape.Copy();
            float shape_width = shape.Width;
            float shape_height = shape.Height;
            float shape_top = shape.Top;
            float shape_left = shape.Left;
            float slide_height = PPTObject.PageSetup.SlideHeight;
            float slide_width = PPTObject.PageSetup.SlideWidth;
            float shape_width_total=0;
            float shape_height_total = 0;
            for (int index = 0; index < Model.Count; index++)
            {
                tbl.Cell(1, 1).Shape.Fill.ForeColor.RGB = Model[index].header_Color;
                Microsoft.Office.Interop.PowerPoint.Cell cell_ = tbl.Cell(1,1);
               // tbl.Background.Fill.BackColor.RGB = 0;
                Microsoft.Office.Interop.PowerPoint.TextFrame tf = cell_.Shape.TextFrame;
                tf.TextRange.Text = Model[index].Header.Name;
                tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Name = Model[index].Header.Name_CFont.FontName;
                tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Color.RGB = Model[index].Header.Name_CFont.Color;
                tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Size = Model[index].Header.Name_CFont.FontSize;

                for (int innerindex = 0; innerindex < Model[index].PlayerName.Count; innerindex++)
                {
                    if(innerindex>0)
                    {
                        tbl.Rows.Add();
                    }
                    cell_ = tbl.Cell(innerindex+2, 1);
                    tf = cell_.Shape.TextFrame;
                    tf.TextRange.Text = Model[index].PlayerName[innerindex].Name;
                    tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Name = Model[index].PlayerName[innerindex].Name_CFont.FontName;
                    tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Color.RGB = Model[index].PlayerName[innerindex].Name_CFont.Color;
                    tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Size = Model[index].PlayerName[innerindex].Name_CFont.FontSize;
                }

              shape_width_total =   shape.Width + shape.Left;
              shape_height_total =  shape.Height + shape.Top;
                //have to check cond

                if(shape_height_total > slide_height  && (shape_width_total + shape.Width + 5.0 ) > slide_width  )
                {
                    shape.Delete();
                    AddSlide(PPTObject, i);
                    i++;
                    sl0 = SL[i];
                    shape = GetShape(SL[i], "Table");
                    tbl = shape.Table;
                    shape.Copy();
                    shape_width = shape.Width;
                    shape_height = shape.Height;
                    shape_top = shape.Top;
                     shape_left = shape.Left;
                    slide_height = PPTObject.PageSetup.SlideHeight;
                     slide_width = PPTObject.PageSetup.SlideWidth;
                    shape_width_total = 0;
                    shape_height_total = 0;
                   
                    index--;
                    
                }
                else if (shape_height_total > slide_height && !((shape_width_total + shape.Width + 5.0) > slide_width))
                {
                    shape.Left = shape_width_total + 5.0F;
                    shape.Top = shape_top;
                    shape_height_total = shape.Height + shape.Top;
                    if (index != Model.Count - 1)
                    {
                        Microsoft.Office.Interop.PowerPoint.ShapeRange shaperange = sl0.Shapes.Paste();
                        shape = shaperange[1];
                        tbl = shape.Table;
                        shape.Left = shape_width_total + 5.0F;
                        shape.Top = shape_height_total + 5.0F;
                    }
                }
                else
                {
                    shape_height_total = shape.Height + shape.Top;
                    float left = shape.Left;
                    if (index != Model.Count - 1)
                    {
                        Microsoft.Office.Interop.PowerPoint.ShapeRange shaperange = sl0.Shapes.Paste();
                        shape = shaperange[1];
                        tbl = shape.Table;
                        shape.Left = left;
                        shape.Top = shape_height_total + 5.0F;
                    }
                }
                 
             
                
                //end
            }

        }
        Microsoft.Office.Interop.PowerPoint.Shape GetShape(Slide slide, string reqobj)
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

    }
}
