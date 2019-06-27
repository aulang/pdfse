using iText.IO.Image;
using Microsoft.Extensions.Options;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Pkcs;
using Org.BouncyCastle.X509;
using PDFService.Models;
using System.IO;

namespace PDFService.Business
{
    public class StamperManager
    {
        private readonly string outBaseDir;
        private readonly ImageData stamper;
        private readonly ICipherParameters pk;
        private readonly X509Certificate[] chain;

        private readonly FileLoaction fileLoaction;
        private readonly CertConfig certConfig;

        public StamperManager(IOptions<FileLoaction> fileOptions, IOptions<CertConfig> certOption)
        {
            this.fileLoaction = fileOptions.Value;
            this.certConfig = certOption.Value;

            using (FileStream fs = new FileStream(fileLoaction.Cert, FileMode.Open))
            {
                stamper = ImageDataFactory.Create(fileLoaction.Stamper);

                Pkcs12Store store = new Pkcs12Store(fs, certConfig.Password.ToCharArray());

                X509CertificateEntry[] entries = store.GetCertificateChain(certConfig.Alias);

                int length = entries.Length;
                chain = new X509Certificate[entries.Length];
                for (int i = 0; i != length; ++i)
                {
                    chain[i] = entries[i].Certificate;
                }

                pk = store.GetKey(certConfig.Alias).Key;
            }

            this.outBaseDir = fileLoaction.OutBaseDir;
        }

        public ImageData Stamper { get => stamper; }
        public string OutBaseDir { get => outBaseDir; }
        public ICipherParameters PrivateKey { get => pk; }
        public X509Certificate[] Chain { get => chain; }
    }
}
