using iText.Forms;
using iText.IO.Image;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Signatures;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.X509;
using System.Collections.Generic;
using System.IO;

namespace PDFService.Utils
{
    public class PdfUtil
    {
        public static void Sign(string input, string output, ImageData stamper, ICipherParameters privateKey, X509Certificate[] chain, string flag)
        {
            PdfDocument document = new PdfDocument(new PdfReader(input));

            PdfAcroForm acroForm = PdfAcroForm.GetAcroForm(document, false);
            bool append = (acroForm != null && acroForm.GetSignatureFlags() != 0);

            int pageNumber = document.GetNumberOfPages();

            RegexBasedLocationExtractionStrategy strategy = new RegexBasedLocationExtractionStrategy(flag);
            PdfDocumentContentParser parser = new PdfDocumentContentParser(document);
            parser.ProcessContent(pageNumber, strategy);
            var locations = new List<IPdfTextLocation>(strategy.GetResultantLocations());

            document.Close();

            PdfSigner signer = new PdfSigner(new PdfReader(input), new FileStream(output, FileMode.Create), append);
            signer.SetCertificationLevel(PdfSigner.CERTIFIED_NO_CHANGES_ALLOWED);

            PdfSignatureAppearance appearance = signer.GetSignatureAppearance();
            appearance.SetPageNumber(pageNumber);

            int size = locations.Count;
            if (size != 0)
            {
                IPdfTextLocation location = locations[size - 1];

                float flagX = location.GetRectangle().GetX();
                float flagY = location.GetRectangle().GetY();

                float width = stamper.GetWidth();
                float height = stamper.GetHeight();

                float x = flagX - width / 2;
                float y = flagY - height / 2;

                appearance.SetRenderingMode(PdfSignatureAppearance.RenderingMode.GRAPHIC);
                appearance.SetSignatureGraphic(stamper);
                appearance.SetPageRect(new Rectangle(x, y, width, height));
            }

            PrivateKeySignature signature = new PrivateKeySignature(privateKey, DigestAlgorithms.SHA256);
            signer.SignDetached(signature, chain, null, null, null, 0, PdfSigner.CryptoStandard.CADES);
        }
    }
}
