namespace PDFService.Models
{
    public class FileLoaction
    {
        private string outBaseDir;
        private string stamper;
        private string cert;

        public string OutBaseDir { get => outBaseDir; set => outBaseDir = value; }
        public string Stamper { get => stamper; set => stamper = value; }
        public string Cert { get => cert; set => cert = value; }
    }
}
