using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using TracxnSlideGenerator;

namespace InvokerApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            if(args.Count()!=3)
            {
                Console.WriteLine("Please Check Your Input Arguments");
            }
            if(FIlecheckForPowerPoint(args)==0)
            {
                return;
            }
            if (FIlecheckForJsonData(args) == 0)
            {
                return;
            }
            if (args[2].Equals("Type1"))
            {
                Initiator(args[0], args[1], args[2]);
            }
            else if (args[2].Equals("Type2"))
            {
                Initiator(args[0], args[1], args[2]);
            }
            else
            {
                Console.WriteLine("we are not currently supported your requested type");
                return;
            }
        }
     
        private static int  FIlecheckForPowerPoint(string[] args)
        {
            string PowerPointFIlePath = args[0];
            int check = Filecheck(PowerPointFIlePath);
            if (check != 0)
            {
                if (check == 1)
                {
                    Console.WriteLine("Power Point file already exists");
                    return 0;
                }
                if (check == 2)
                {
                    Console.WriteLine("Please check Power Point file Path / Path format");
                    return 0;
                }
            }
            return 1;
        }
        private static int FIlecheckForJsonData(string[] args)
        {
            string JsonFIlePath = args[1];
            int check = Filecheck(JsonFIlePath);
            if (check != 1)
            {
                if (check == 0)
                {
                    Console.WriteLine("Json FIle not exists");
                    return 0;
                }
                if (check == 2)
                {
                    Console.WriteLine("Please check Json file Path / Path format");
                    return 0;
                }
            }
            return 1;
        }

        static int Filecheck(string Path)
        {
            int i = 0;
            try
            {
              bool check =  File.Exists(@Path);
              if (check)
              {
                  i = 1;
              }
              else
              {
                  i = 0;
              }
            }
            catch
            {
                i = 2;
            }
            return i;
        }

        static bool flag = false;
        static bool Check = false;
        private static void Initiator(string PPTFilePath, string Jsonfilepath, string Type, bool REC = false)
        {
            bool update = false;
            if(REC==true)
            {
                update = true;
                if(flag)
                {
                    Console.WriteLine("file is currently in use wait ... ");
                    Check = true;
                }
                else
                {
                    try
                    {
                        File.Delete(PPTFilePath);
                    }
                    catch
                    {
                        Console.WriteLine("file is currently in use ... ");
                    }
                    
                    if(Check)
                    {
                        Check = false;
                        Initiator(PPTFilePath, Jsonfilepath, Type, true);
                    }
                }
            }
            flag = true;
            GenratePPt(PPTFilePath, Jsonfilepath,Type,update);

            if (!REC)
            {
                Console.WriteLine("process completed Please Check Your FIle");
                flag = false;
                CheckFOrChange(PPTFilePath, Jsonfilepath, Type);
            }

            flag = false;
            Console.WriteLine("process completed Please Check Your FIle");
        }


        private static void GenratePPt(string PPTFilePath, string Jsonfilepath,string Type,bool Update)
        {
            Generator gen = new Generator();
            string json;
            using (StreamReader r = new StreamReader(Jsonfilepath))
            {
                json = r.ReadToEnd();
            }
            gen.BuildSlide(Type, PPTFilePath, json,Update);
        }


    

        private static string path_ = "";
        private static string JsonFilePath_ = "";
        private static string type_ = "";
        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
        public static void CheckFOrChange(string path, string JsonFilePath,string type)
        {
            path_ = path;
            JsonFilePath_ = JsonFilePath;
            type_ = type;
            FileSystemWatcher m_Watcher = new FileSystemWatcher();
            m_Watcher.Path = Path.GetDirectoryName(JsonFilePath);
            m_Watcher.Filter = Path.GetFileName(JsonFilePath);
            m_Watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite;
            m_Watcher.Changed += new FileSystemEventHandler(OnChanged);
            m_Watcher.EnableRaisingEvents = true;
            Console.WriteLine("Press \'q\' to quit Sync Process.");
            while (Console.Read() != 'q') ;
        }
        private static void OnChanged(object source, FileSystemEventArgs e)
        {
          //  Console.WriteLine("File changed");
            FileInfo info = new FileInfo(JsonFilePath_);
            bool res = IsFileLocked(info);
            if (!res)
            {
                Console.WriteLine("File changed, we are updating your file please do not open your file. We will inform you when we are done , and please try not to update Json File until we done with processing your current updates ");
                Initiator(path_, JsonFilePath_, type_, true);
            }
        }
        static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
            
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
    }
}
