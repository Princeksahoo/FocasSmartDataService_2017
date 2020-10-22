using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;
using System.Threading;

namespace FocasSmartDataCollection
{
    public static class CleanUpProcess
    {
        static string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        public static void DeleteFiles(string folder, string ip)
        {           
            DirectoryInfo di = new DirectoryInfo(Path.Combine(appPath,folder));
            FileInfo[] files = di.GetFiles("*.txt");

            if (files.Length > 0)
            {
                int daysForDeleteFile = DatabaseAccess.GetLoghistorydays();
                foreach (FileInfo fi in files)
                {                    
                    TimeSpan ts = DateTime.Now.Subtract(fi.LastWriteTime);
                    int days = ts.Days + 1;                    
                    if (days >= daysForDeleteFile)
                    {
                        try
                        {
                            fi.Delete();
                        }
                        catch { }
                    }
                }
            }
        }

        public static void RenameLogFiles()
        {
            string progTime = String.Format("_{0:yyyyMMdd}", DateTime.Now);
            string location = appPath + "\\Logs\\F-" + Thread.CurrentThread.Name + progTime + "-Status.txt";
            FileInfo f = new FileInfo(location);
            if (f.Exists && f.Length > 1048576)
            {
                string newfile = appPath + "\\Logs\\" + (Path.GetFileNameWithoutExtension(f.Name)) + String.Format("_{0:HHmmss}", DateTime.Now) + ".txt";// + String.Format("{0:HHmmss}", DateTime.Now));
                try
                {
                    f.MoveTo(newfile);
                    Thread.Sleep(10);
                    Logger.WriteDebugLog( string.Format("File {0} has been renamed to {1}.", location, newfile));
                }
                catch(Exception ex)
                {
                    Logger.WriteErrorLog(" RenameLogFiles(): " + ex.Message);
                }               
            }
        }
        public static void RenameTPMFiles()
        {
            string progTime = String.Format("_{0:yyyyMMdd}", DateTime.Now);
            string location = appPath + "\\TPMFiles\\F-" + Thread.CurrentThread.Name + progTime + ".txt";
            
            FileInfo f = new FileInfo(location);
            if (f.Exists && f.Length > 1048576)
            {
                string newfile = appPath + "\\TPMFiles\\" + (Path.GetFileNameWithoutExtension(f.Name)) + String.Format("_{0:HHmmss}", DateTime.Now) + ".txt";
                try
                {
                    f.MoveTo(newfile);
                    Logger.WriteDebugLog(string.Format("File {0} has been renamed to {1}.", location, newfile));
                }
                catch (Exception ex)
                {
                    Logger.WriteErrorLog(" RenameTPMFiles(): " + ex.Message);
                }
            }
        }
    }
}
