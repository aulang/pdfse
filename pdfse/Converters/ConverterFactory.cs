using System;

namespace PDFService.Converters
{
    enum DocumentType
    {
        DOC,
        DOCX,
        XLS,
        XLSX,
        PPT,
        PPTX
    }

    public class ConverterFactory
    {
        public const string DOT = ".";

        public static Converter GetConverter(string extension)
        {
            try
            {
                Converter converter = null;
                if (extension.StartsWith(DOT))
                {
                    extension = extension.Remove(0, 1).ToUpper();
                }

                DocumentType type = (DocumentType)Enum.Parse(typeof(DocumentType), extension);
                switch (type)
                {
                    case DocumentType.DOC:
                        converter = new WordConverter();
                        break;
                    case DocumentType.DOCX:
                        converter = new WordConverter();
                        break;
                    case DocumentType.XLS:
                        converter = new ExcelConverter();
                        break;
                    case DocumentType.XLSX:
                        converter = new ExcelConverter();
                        break;
                    case DocumentType.PPT:
                        converter = new PowerPointConverter();
                        break;
                    case DocumentType.PPTX:
                        converter = new PowerPointConverter();
                        break;
                }
                return converter;
            }
            catch
            {
                throw new ConvertException("Unsupported file format");
            }
        }
    }
}
