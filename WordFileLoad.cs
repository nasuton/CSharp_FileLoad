using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace CSharp_FileLoad
{
    class WordFileLoad
    {
        public WordFileLoad()
        {
            
        }

        private void WordRead(string _filePath)
        {
            var word = new Word.Application();
            try
            {
                //Wordを非表示で実行
                word.Visible = false;
                //警告ウィンドウを表示しない
                word.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                //マクロを無効化
                word.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                Word.Documents docs = word.Documents;
                try
                {
                    //Wordを開くときのパスワード(必要ない場合は無視される)
                    string openPass = "a";
                    //テンプレートを開くためのパスワード(必要ない場合は無視される)
                    string tempOpenPass = "a";
                    //文書への変更を保存するためのパスワード(必要ない場合は無視される)
                    string writePass = "a";
                    //テンプレートへの変更を保存するためのパスワード(必要ない場合は無視される)
                    string tempWritePass = "a";
                    var doc = docs.Open(_filePath, Type.Missing, Type.Missing, Type.Missing, openPass, tempOpenPass, Type.Missing, writePass, tempWritePass);
                    try
                    {
                        //Wordファイル内に設定されているハイパーリンクを取得する
                        foreach(Word.Hyperlink hyperlink in doc.Hyperlinks)
                        {
                            Console.WriteLine(hyperlink.Address);
                        }

                        //行ごとに読み込み
                        foreach(Word.Paragraph paragraph in doc.Paragraphs)
                        {
                            Console.WriteLine(paragraph.Range.Text);
                        }

                        //図形にあるテキストボックスから取得
                        foreach(Word.Shape shape in doc.Shapes)
                        {
                            //描画キャンバス
                            if(shape.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
                            {
                                //キャンバス内すべてのアイテムに対して実行
                                foreach(Word.Shape item in shape.CanvasItems)
                                {
                                    if(item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                    {
                                        foreach(Word.Shape groupItem in item.GroupItems)
                                        {
                                            if(groupItem.TextFrame.HasText == -1)
                                            {
                                                Console.WriteLine(groupItem.TextFrame.TextRange.Text);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if(item.TextFrame.HasText == -1)
                                        {
                                            Console.WriteLine(item.TextFrame.TextRange.Text);
                                        }
                                    }
                                }
                            }
                            //グループ化されていた場合
                            else if(shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                            {
                                foreach(Word.Shape item in shape.GroupItems)
                                {
                                    if(item.TextFrame.HasText == -1)
                                    {
                                        Console.WriteLine(item.TextFrame.TextRange.Text);
                                    }
                                }
                            }
                            //それ以外
                            else
                            {
                                if(shape.TextFrame.HasText == -1)
                                {
                                    Console.WriteLine(shape.TextFrame.TextRange.Text);
                                }
                            }
                        }
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                    finally
                    {
                        if(doc != null)
                        {
                            //Wordを保存せずに閉じる
                            doc.Close(false);
                            Marshal.ReleaseComObject(doc);
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                finally
                {
                    Marshal.ReleaseComObject(docs);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                if (word != null)
                {
                    word.Quit();
                    Marshal.ReleaseComObject(word);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
            }
        }

        public void FileRead(string _filePath)
        {
            //対象ファイルが存在するか確認する
            if (File.Exists(_filePath))
            {
                //ファイルの拡張子がWordかどうかチェック
                string ext = Path.GetExtension(_filePath);
                Match match = Regex.Match(ext, ".(docx?|docm)");
                if (match.Value != "")
                {
                    WordRead(_filePath);
                }
                else
                {
                    Console.WriteLine("対象となるファイルはWordではありませんでした。");
                }
            }
            else
            {
                Console.WriteLine("対象となるファイルが存在しませんでした。");
            }
        }
    }
}
