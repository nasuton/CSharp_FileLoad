using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSharp_FileLoad
{
    class ExcelFileLoad
    {
        public ExcelFileLoad()
        {

        }


        private void ExcelRead(string _filePath)
        {
            Excel.Application excel = new Excel.Application();
            try
            {
                //Excelを非表示で実行
                excel.Visible = false;
                //警告ウィンドウを表示しない
                excel.DisplayAlerts = false;
                //マクロを実行させない
                excel.EnableEvents = false;
                excel.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                Excel.Workbooks workBooks = excel.Workbooks;
                try
                {
                    //Excelを開く際にパスワードが必要な場合(必要ない場合は無視される)
                    string openPass = "a";
                    //Excelを書き込む際にパスワードが必要な場合(必要ない場合は無視される)
                    string writePass = "a";
                    //Excelを開く
                    Excel.Workbook workBook = workBooks.Open(_filePath, Type.Missing, Type.Missing, Type.Missing, openPass, writePass);
                    try
                    {
                        Excel.Sheets workSheets = workBook.Sheets;
                        try
                        {
                            foreach(Excel.Worksheet sheet in workSheets)
                            {
                                try
                                {
                                    //Excelシートに設定されているハイパーリンクを取得する
                                    foreach(Excel.Hyperlink hyperlink in sheet.Hyperlinks)
                                    {
                                        Console.WriteLine(hyperlink.Address);
                                    }

                                    //Excelシートの図形のテキストボックスから取得する
                                    foreach(Excel.Shape shape in sheet.Shapes)
                                    {
                                        //図形がグループ化されている場合
                                        if(shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                        {
                                            foreach(Excel.GroupShapes groupItem in shape.GroupItems)
                                            {
                                                foreach(Excel.Shape sha in groupItem)
                                                {
                                                    if(sha.TextFrame2.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                                                    {
                                                        Console.WriteLine(sha.TextFrame2.TextRange.Text);
                                                    }
                                                }
                                            }
                                        }
                                        //それ以外の場合
                                        else
                                        {
                                            if (shape.TextFrame2.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                                            {
                                                Console.WriteLine(shape.TextFrame2.TextRange.Text);
                                            }
                                        }
                                    }

                                    //行を読み込み
                                    foreach(Excel.Range row in sheet.UsedRange.Rows)
                                    {
                                        //列を読み込み
                                        foreach(Excel.Range cln in row.Columns)
                                        {
                                            Console.WriteLine(cln.Text);
                                        }
                                    }

                                }
                                catch(Exception ex)
                                {
                                    Console.WriteLine(ex.ToString());
                                }
                                finally
                                {
                                    Marshal.ReleaseComObject(sheet);
                                }
                            }
                        }
                        catch(Exception ex)
                        {
                            Console.WriteLine(ex.ToString());
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(workSheets);
                        }
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                    finally
                    {
                        if(workBook != null)
                        {
                            //Excelを保存せずに閉じる
                            workBook.Close(false);
                            Marshal.ReleaseComObject(workBook);
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                    }
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                finally
                {
                    Marshal.ReleaseComObject(workBooks);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                if(excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
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
                //ファイルの拡張子がExcelかどうかチェック
                string ext = Path.GetExtension(_filePath);
                Match match = Regex.Match(ext, ".(xlsx?|xlsm)");
                if(match.Value != "")
                {
                    ExcelRead(_filePath);
                }
                else
                {
                    Console.WriteLine("対象となるファイルはExcelではありませんでした。");
                }
            }
            else
            {
                Console.WriteLine("対象となるファイルが存在しませんでした。");
            }
        }


    }
}
