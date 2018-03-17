using System;
using System.IO;
using System.Text;

namespace PDFService.Converters
{
    public abstract class Converter
    {
        public abstract void Convert(String inputFile, String outputFile);

        protected void ReleaseCOMObject(object obj)
        {
            try
            {
                if (System.Runtime.InteropServices.Marshal.IsComObject(obj))
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
                }
            }
            catch
            {
            }
            obj = null;
        }
    }
}
