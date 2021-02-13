using System;

namespace CSharp_FileLoad
{
    class Program
    {
        static void Main(string[] args)
        {
            //Excelファイル読み込み
            ExcelFileLoad excelRead = new ExcelFileLoad();
            excelRead.FileRead(@"Excelファイルまでのフルパス");
            //Wordファイル読み込み
            WordFileLoad wordRead = new WordFileLoad();
            wordRead.FileRead(@"Wordファイルまでのフルパス");
            //Pdfファイル読み込み
            PdfFileLoad pdfRead = new PdfFileLoad();
            pdfRead.FileLoad(@"Pdfファイルまでのフルパス");
            //Txtファイルを読み込み
            TxtFileLoad txtRead = new TxtFileLoad();
            txtRead.FileRead(@"テキストファイルまでのフルパス");

            Console.WriteLine();
            Console.WriteLine("終了するには何か押してください。");
            Console.ReadLine();
        }
    }
}
