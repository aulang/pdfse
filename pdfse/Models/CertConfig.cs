namespace PDFService.Models
{
    public class CertConfig
    {
        private string alias;
        private string password;

        public string Alias { get => alias; set => alias = value; }
        public string Password { get => password; set => password = value; }
    }
}
