using System;
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
                if (args.Length == 3)
                {
                    Stopwatch watch = new Stopwatch();
                    watch.Start();
                    if (args[0].Equals("word"))
                        WordConverter.Convert(args[1], args[2]);
                    else if (args[0].Equals("ppt"))
                        PowerpointConverter.Convert(args[1], args[2]);
                    else
                    {
                        watch.Stop();
                        throw new ArgumentException("Invalid type.");
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
    }
}