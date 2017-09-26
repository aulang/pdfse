using iText.Forms;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Signatures;
using System.Collections.Generic;
using System.IO;

namespace PDFService.Utils
{
    public class PdfUtil
    {
        public static void Sign(string flag) {
            PdfDocument document = new PdfDocument(new PdfReader(""));

            PdfAcroForm acroForm = PdfAcroForm.GetAcroForm(document, false);
            bool append = acroForm.GetSignatureFlags() != 0;

            int pageNumber = document.GetNumberOfPages();

            RegexBasedLocationExtractionStrategy strategy = new RegexBasedLocationExtractionStrategy(flag);
            PdfDocumentContentParser parser = new PdfDocumentContentParser(document);
            parser.ProcessContent(pageNumber, strategy);
            var locations = new List<IPdfTextLocation>(strategy.GetResultantLocations());

            document.Close();

            PdfSigner signer = new PdfSigner(new PdfReader(""), new FileStream("", FileMode.Create), append);

        }
    }
}
