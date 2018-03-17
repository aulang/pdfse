using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace PDFService.Converters
{
    public class WordConverter : Converter
    {
        private Word.Application app;
        private Word.Document doc;

        public override void Convert(String inputFile, String outputFile)
        {
            Object nothing = System.Reflection.Missing.Value;
            try
            {
                if (!File.Exists(inputFile))
                {
                    throw new ConvertException("File not Exists");
                }

                app = new Word.Application();
                doc = app.Documents.Open(inputFile, false, true, false, nothing, nothing, true, nothing, nothing, nothing, nothing, false, false, nothing, true, nothing);
                doc.ExportAsFixedFormat(outputFile, Word.WdExportFormat.wdExportFormatPDF, false, Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen,
                    Word.WdExportRange.wdExportAllDocument, 1, 1, Word.WdExportItem.wdExportDocumentContent, false, false, Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                    false, false, false, nothing);
            }
            catch (Exception e)
            {
                throw new ConvertException(e.Message);
            }
            finally {
                Release();
            }
        }

        private void Release()
        {
            if (doc != null)
            {
                try
                {
                    doc.Close(false);
                    ReleaseCOMObject(doc);
                }
                catch
                {                    
                }
            }
            
            if (app != null)
            {
                try
                {
                    app.Quit();
                    ReleaseCOMObject(app);
                }
                catch
                {
                }
            }
        }
    }
}
