using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace druckmessungsprotokoll
{
    internal class Logger
    {
        private string CurrentDirectory { get; set; }
        private string FileName { get; set; }  
        private string FilePath { get; set; }

        public Logger()
        {
            this.CurrentDirectory = Directory.GetCurrentDirectory();
            this.FileName = "Log.txt";
            this.FilePath = Path.Combine(CurrentDirectory, this.FileName);
        }

        public void Log(string message)
        {
            using(StreamWriter sw = File.AppendText(this.FilePath))
            {
                Console.WriteLine(message);
                sw.WriteLine($"Log from {DateTime.Now.ToLongTimeString}, {DateTime.Now.ToShortDateString} - {message}");
            }
        }

    }
}
