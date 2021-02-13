using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace CSharp_FileLoad
{
    class PowerPointFileLoad
    {
        public PowerPointFileLoad()
        {

        }

        private void PowerPointRead(string _filePath)
        {
            PowerPoint.Application powerPoint = new PowerPoint.Application();
            try
            {
                //PowerPointを非表示で実行
                powerPoint.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                //警告ウィンドウを表示しない
                powerPoint.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                if(powerPoint != null)
                {
                    powerPoint.Quit();
                    Marshal.ReleaseComObject(powerPoint);
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
                //ファイルの拡張子がPowerPointかどうかチェック
                string ext = Path.GetExtension(_filePath);
                Match match = Regex.Match(ext, ".(pptx?|pptm)");
                if (match.Value != "")
                {
                    PowerPointRead(_filePath);
                }
                else
                {
                    Console.WriteLine("対象となるファイルはPowerPointではありませんでした。");
                }
            }
            else
            {
                Console.WriteLine("対象となるファイルが存在しませんでした。");
            }
        }
    }
}
