using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace TestAutoGenerator
{
    public static class Logger
    {
        public static string LogPath = string.Empty;
        public static string ExecutionTimeString = string.Empty;

        static Logger()
        {
            var m_exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            ExecutionTimeString = DateTime.Now.ToString("yyyy-MM-d_HH-mm-ss");
            LogPath = m_exePath + "\\" + "log_" + ExecutionTimeString + ".txt";
        }

        public static void Log(string message)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(LogPath, true))
                {
                    sw.Write("\r\nLog Entry : ");
                    sw.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                    sw.WriteLine("  :{0}", message);
                    sw.WriteLine("---------------------------------------------------------------------------");
                }
            }
            catch (Exception e)
            {}
        }

        public static void Log(Exception ex)
        {
            try
            {
                using (StreamWriter sw = File.AppendText(LogPath))
                {
                    sw.Write("\r\nLog Entry : ");
                    sw.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                    sw.WriteLine("Application exception: ");
                    sw.WriteLine("Error: {0}", ex.Message);
                    sw.WriteLine("StackTrace: {0}", ex.StackTrace);
                    sw.WriteLine("---------------------------------------------------------------------------");
                }               
            }
            catch (Exception e)
            {}
        }

        public static void Log(Exception ex, string message)
        {
            try
            {
                using (StreamWriter sw = File.AppendText(LogPath))
                {
                    sw.Write("\r\nLog Entry : ");
                    sw.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                    sw.WriteLine("  :{0}", message);
                    sw.WriteLine("Error: {0}", ex.Message);
                    sw.WriteLine("StackTrace: {0}", ex.StackTrace);
                    sw.WriteLine("---------------------------------------------------------------------------");
                }
            }
            catch (Exception e)
            { }
        }
    }
}
