using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace CSharp_FileLoad
{
    class PdfFileLoad
    {
        public PdfFileLoad()
        {

        }

        private void PdfRead(string _filePath)
        {
            try
            {
                //Pdfファイルを開く。パスワードがかかっていた場合はエラーとなる
                PdfReader reader = new PdfReader(_filePath);
                //Pdfファイルのページ数を取得
                int pages = reader.NumberOfPages;
                for(int i = 1; i <= pages; i++)
                {
                    //1ページずつ読み込む
                    string text = PdfTextExtractor.GetTextFromPage(reader, i);
                    var lines = text.Replace("\r\n", "\n").Split(new[] { '\n', '\r' });
                    foreach(string line in lines)
                    {
                        Console.WriteLine(line);
                    }

                    reader.Close();
                    reader.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public void FileLoad(string _filePath)
        {
            if (File.Exists(_filePath))
            {
                string ext = System.IO.Path.GetExtension(_filePath);
                Match match = Regex.Match(ext, ".pdf");
                if(match.Value != "")
                {
                    PdfRead(_filePath);
                }
                else
                {
                    Console.WriteLine("対象となるファイルはPDFではありませんでした。");
                }
            }
            else
            {
                Console.WriteLine("対象となるファイルが存在しませんでした。");
            }
        }
    }
}
