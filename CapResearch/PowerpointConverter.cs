using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace CapResearch
{
    public static class PowerpointConverter
    {
        public static void Convert(string input, string output)
        {
            PowerPoint.Application pptApp = null;
            try
            {
                pptApp = new PowerPoint.Application();
                var powerpointDocument = pptApp.Presentations.Open(input,
                    MsoTriState.msoTrue, //ReadOnly
                    MsoTriState.msoFalse, //Untitled
                    MsoTriState.msoFalse); //Window not visible during converting
                powerpointDocument.ExportAsFixedFormat(output,
                    PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
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