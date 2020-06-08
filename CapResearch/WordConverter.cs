using System;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace CapResearch
{
    public static class WordConverter
    {
        public static void Convert(string input, string output)
        {
            Word.Application wordApp = null;
            try
            {
                wordApp = new Word.Application
                {
                    Visible = false
                };

                var wordDoc = wordApp.Documents.Open(input, ReadOnly: true); // Open in readonly
                var viewQuality = Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen;
                wordDoc.ExportAsFixedFormat(output, Word.WdExportFormat.wdExportFormatPDF, false, viewQuality);
                wordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges,
                    Word.WdOriginalFormat.wdOriginalDocumentFormat,
                    false); //Close document
            }
            catch (COMException e)
            {
                Console.WriteLine("Microsoft Word is not installed...");
                Console.WriteLine(e.Message);
            }
            finally
            {
                wordApp?.Quit();  //Important: When you forget this Word keeps running in the background
            }
        }
    }
}