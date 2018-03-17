using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace PDFService.Converters
{
    public class ExcelConverter : Converter
    {
        private Excel.Application app;
        private Excel.Workbook book;

        public override void Convert(String inputFile, String outputFile)
        {
            Object nothing = Type.Missing;
            try
            {
                if (!File.Exists(inputFile))
                {
                    throw new ConvertException("File not Exists");
                }

                app = new Excel.Application();
                book = app.Workbooks.Open(inputFile, false, true, nothing, nothing, nothing, true, nothing, nothing, false, false, nothing, false, nothing, false);

                bool hasContent = false;
                foreach (Excel.Worksheet sheet in book.Worksheets)
                {
                    Excel.Range range = sheet.UsedRange;
                    if (range != null)
                    {
                        Excel.Range found = range.Cells.Find("*", nothing, nothing, nothing, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, nothing, nothing, nothing);
                        if (found != null)
                        {
                            hasContent = true;
                        }
                        ReleaseCOMObject(found);
                        ReleaseCOMObject(range);
                    }
                }

                if (!hasContent)
                {
                    throw new ConvertException("No Content");
                }
                book.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputFile, Excel.XlFixedFormatQuality.xlQualityMinimum, false, false, nothing, nothing, false, nothing);
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
            if (book != null)
            {
                try
                {
                    book.Close(false);
                    ReleaseCOMObject(book);
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
