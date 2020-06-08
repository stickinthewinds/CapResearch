using System;
using System.IO;
using System.Diagnostics;

namespace CapResearch
{
  internal class Program
    {
        /// <summary>
        /// Take a type, input file and output file from the command line in that order and attempt to convert the file to pdf
        /// </summary>
        /// <param name="args"></param>
        public static void Main(string[] args)
        {
            try
            {
                if (args.Length >= 2)
                {
                    var watch = new Stopwatch();
                    watch.Start();

                    switch (GetExtension(args[0]))
                    {
                        case "docx": case "doc": case "odt":
                            WordConverter.Convert(args[0], args[1], GetExtension(args[1]));
                            break;

                        case "pptx": case "ppt":
                            PowerpointConverter.Convert(args[0], args[1], GetExtension(args[1]));
                            break;

                        default:
                            watch.Stop();
                            throw new ArgumentException("Invalid file type");
                    }

                    watch.Stop();
                    Console.WriteLine($"Time taken for conversion of {args[0]} to PDF: {watch.ElapsedMilliseconds}ms");
                }
                else
                {
                    throw new ArgumentException("Invalid arguments.\n" +
                                                "Arguments are 'FileType' 'input file' 'output file'." +
                                                "\nFileType can be either 'word' or 'ppt'." +
                                                "\nExample input file for word: C:\\Users\\Example\\example.docx" +
                                                "\nExample output file: C:\\Users\\Example\\example.pdf");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public static string GetExtension(string path) => Path.GetExtension(path).Substring(1).ToLower();
    }
}