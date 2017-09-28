using System;

namespace PDFService.Converters
{
    public class ConvertException : Exception
    {
        public ConvertException(String message) : base(message)
        {
        }
    }
}
