using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace CapResearch
{
    public static class PowerpointConverter
    {
        public static void Convert(string input, string output, string type)
        {
            PowerPoint.Application pptApp = null;
            try
            {
                PowerPoint.PpFixedFormatType format;
                if (type.Equals("xps"))
                    format = PowerPoint.PpFixedFormatType.ppFixedFormatTypeXPS;
                else if (type.Equals("pdf"))
                    format = PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF;
                else
                    throw new ArgumentException("Invalid output file type");

                pptApp = new PowerPoint.Application();
                var powerpointDocument = pptApp.Presentations.Open(input,
                    MsoTriState.msoTrue, //ReadOnly
                    MsoTriState.msoFalse, //Untitled
                    MsoTriState.msoFalse); //Window not visible during converting
                powerpointDocument.ExportAsFixedFormat(output, format);
                powerpointDocument.Close(); //Close document
            }
            catch (COMException e)
            {
                Console.WriteLine("Microsoft PowerPoint is not installed...");
                Console.WriteLine(e.Message);
            }
            finally
            {
                pptApp?.Quit(); //Important: When you forget this PowerPoint keeps running in the background
            }
        }
    }
}