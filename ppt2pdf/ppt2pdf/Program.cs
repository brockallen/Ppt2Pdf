using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Linq;
using System.IO;

namespace ppt2pdf
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                PrintUsage();
                return;
            }

            var path = args[0];
            if (Directory.Exists(path))
            {
                Run(path);
            }
            else
            {
                Console.WriteLine("{0} is not a directory");
            }
        }

        private static void PrintUsage()
        {
            Console.WriteLine("Pass the directory as a param");
        }

        private static void Run(string path)
        {
            var files = Directory.GetFiles(path, "*.pptx").Union(Directory.GetFiles(path, "*.ppt"));
            foreach (var file in files)
            {
                if (!Path.GetFileName(file).StartsWith("~"))
                {
                    Process(file);
                }
            }

            var subs = Directory.GetDirectories(path);
            foreach (var sub in subs)
            {
                Run(sub);
            }
        }

        private static void Process(string file)
        {
            Console.WriteLine("Found: {0}", file);
            File.SetAttributes(file, FileAttributes.Normal);
            var idx = file.LastIndexOf(".");
            var newPath = file.Remove(idx, file.Length - idx) + ".pdf";
            Convert(file, newPath);
        }

        static void Convert(string path, string newPath)
        {
            var a = new Application();
            var p = a.Presentations.Open(path);
            //var newName = path.Replace(".pptx", ".pdf");
            p.SaveAs(newPath, PpSaveAsFileType.ppSaveAsPDF);
            p.Close();
        }
    }
}
