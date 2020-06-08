using System;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace CapResearch
{
    public static class WordConverter
    {
        public static void Convert(string input, string output)
        {
            try
            {
                var wordApp = new Word.Application();
                wordApp.Visible = false;
                
                var wordDoc = wordApp.Documents.Open(input,
                    ReadOnly: true); // Open in readonly
                
                Word.WdExportOptimizeFor viewQuality = Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen;
            
                wordDoc.ExportAsFixedFormat(output, Word.WdExportFormat.wdExportFormatPDF, false, viewQuality);

                wordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges, 
                    Word.WdOriginalFormat.wdOriginalDocumentFormat, 
                    false); //Close document

                wordApp.Quit(); //Important: When you forget this Word keeps running in the background
            }
            catch (COMException e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}