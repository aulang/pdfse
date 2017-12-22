using PDFService.Converters;
using PDFService.Utils;
using System;
using System.IO;
using System.Text;

namespace PDFService.Business
{
    public class FileManager
    {
        public const string PDF = "pdf";
        public const string DOC = "doc";
        public const string SIGNED = "signed";

        private const string DEFAULTFLAG = "盖章";
        private const string DATEFORMAT = "yyyyMMdd";

        private static volatile int count = 0;

        private StamperManager Stamper;

        public FileManager(StamperManager stamper)
        {
            this.Stamper = stamper;
        }

        private string GetParentDirectory(string path)
        {
            string currentDir = Path.GetDirectoryName(path);
            int lastIndex = currentDir.LastIndexOf(Path.DirectorySeparatorChar);

            return currentDir.Remove(lastIndex);
        }

        private string GetInputFilePath(string fileName, string type)
        {
            ++count;

            String today = DateTime.Now.ToString(DATEFORMAT);

            StringBuilder builder = new StringBuilder(Stamper.OutBaseDir);
            builder.Append(Path.DirectorySeparatorChar).Append(today).Append(Path.DirectorySeparatorChar).Append(type);

            string parentDir = builder.ToString();

            if (!Directory.Exists(parentDir))
            {
                count = 0;
                Directory.CreateDirectory(parentDir);
            }


            builder.Append(Path.DirectorySeparatorChar).Append(count).Append(fileName);
            return builder.ToString();
        }

        public string GetInputDocPath(string fileName)
        {
            return GetInputFilePath(fileName, DOC);
        }

        public string GetInputPdfPath(string fileName)
        {
            return GetInputFilePath(fileName, PDF);
        }

        public string GetOutputFilePath(string inputPath)
        {
            StringBuilder builder = new StringBuilder(GetParentDirectory(inputPath));
            builder.Append(Path.DirectorySeparatorChar).Append(PDF);

            string outParentDir = builder.ToString();
            if (!Directory.Exists(outParentDir))
            {
                Directory.CreateDirectory(outParentDir);
            }

            String inputName = Path.GetFileNameWithoutExtension(inputPath);
            builder.Append(Path.DirectorySeparatorChar).Append(Path.ChangeExtension(inputName, PDF));

            return builder.ToString();
        }

        private string GetSignedFilePath(string input)
        {
            string signedDir = GetParentDirectory(input) + Path.DirectorySeparatorChar + SIGNED;

            if (!Directory.Exists(signedDir))
            {
                Directory.CreateDirectory(signedDir);
            }

            return signedDir + Path.DirectorySeparatorChar + Path.GetFileName(input);
        }

        public string Convert(string inputPath)
        {
            string outputPath = GetOutputFilePath(inputPath);

            string extension = Path.GetExtension(inputPath);

            Converter converter = ConverterFactory.GetConverter(extension);

            converter.Convert(inputPath, outputPath);

            return outputPath;
        }

        public string Sign(string input, string flag)
        {
            if (String.IsNullOrWhiteSpace(flag))
            {
                flag = DEFAULTFLAG;
            }

            string signedFilePath = GetSignedFilePath(input);
            PdfUtil.Sign(input, signedFilePath, Stamper.Stamper, Stamper.PrivateKey, Stamper.Chain, flag);
            return signedFilePath;
        }
    }
}
