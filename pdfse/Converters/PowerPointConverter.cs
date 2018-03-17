using System;
using System.IO;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PDFService.Converters
{
    public class PowerPointConverter : Converter
    {
        private PowerPoint.Application app;
        private PowerPoint.Presentation presentation;

        public override void Convert(String inputFile, String outputFile)
        {
            try
            {
                if (!File.Exists(inputFile))
                {
                    throw new ConvertException("File not Exists");
                }

                app = new PowerPoint.Application();
                presentation = app.Presentations.Open(inputFile, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                presentation.ExportAsFixedFormat(outputFile, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
            }
            catch (Exception e)
            {
                throw new ConvertException(e.Message);
            }
            finally
            {
                Release();
            }
        }

        private void Release()
        {
            if (presentation != null)
            {
                try
                {
                    presentation.Close();
                    ReleaseCOMObject(presentation);
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
