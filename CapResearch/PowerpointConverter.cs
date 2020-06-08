using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace CapResearch
{
    public class PowerpointConverter
    {
        public static void Convert(string input, string output)
        {
            var pptApp = new PowerPoint.Application();
            var powerpointDocument = pptApp.Presentations.Open(input,
                MsoTriState.msoTrue, //ReadOnly
                MsoTriState.msoFalse, //Untitled
                MsoTriState.msoFalse); //Window not visible during converting
           
            powerpointDocument.ExportAsFixedFormat(output, 
                PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);

            powerpointDocument.Close(); //Close document
            pptApp.Quit(); //Important: When you forget this PowerPoint keeps running in the background
        }
    }
}