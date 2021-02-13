using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CSharp_FileLoad
{
    class TxtFileLoad
    {
        public TxtFileLoad()
        {

        }

        private void TextRead(string _filePath)
        {
            //対象ファイルを読み込む
            using (StreamReader st = new StreamReader(_filePath, Encoding.GetEncoding("UTF-8")))
            {
                //一行づつ読み込む
                while(st.Peek() != -1)
                {
                    Console.WriteLine(st.ReadLine());
                }
            }
        }

        public void FileRead(string _filePath)
        {
            if (File.Exists(_filePath))
            {
                string ext = Path.GetExtension(_filePath);
                Match match = Regex.Match(ext, ".txt");
                if(match.Value != "")
                {
                    TextRead(_filePath);
                }
                else
                {
                    Console.WriteLine("対象となるファイルはTXTではありませんでした。");
                }
            }
            else
            {
                Console.WriteLine("対象となるファイルが存在しませんでした。");
            }
        }
    }
}
