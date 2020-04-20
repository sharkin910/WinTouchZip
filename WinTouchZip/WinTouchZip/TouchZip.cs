using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WinTouchZip
{
    public static class TouchZip
    {
        public static bool AddYMD { get; set; }

        public static string Exec(params string[] args)
        {
            foreach (var zipPath in args)
            {
                if (zipPath == "-ymd" || zipPath == "/ymd")
                {
                    AddYMD = true;
                }
            }
            if (!AddYMD)
            {
                DialogResult result = MessageBox.Show("ZIPファイル名に年月日を付与しますか？", "確認",
                  MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                  MessageBoxDefaultButton.Button2);
                if (result == DialogResult.Yes)
                {
                    AddYMD = true;
                }
            }

            foreach (var zipPath in args)
            {
                if (zipPath == "-ymd" || zipPath == "/ymd")
                {
                }
                else
                {
                    try
                    {
                        DateTime dt = ZipGetLastModifiedFileTime(zipPath, out string lastModifiedFileName);
                        if (dt != DateTime.MinValue)
                        {
                            string msg = ""
                                + "ZIPファイル名: " + zipPath + "の日付を\r\n"
                                + "ZIP内最新ファイル（" + lastModifiedFileName + "）\r\n"
                                + "のの更新日（" + dt + "）に合わせます。\r\n\r\n"
                                + "実行しますか？";
                            DialogResult result = MessageBox.Show(msg, "実行確認",
                                          MessageBoxButtons.YesNoCancel, MessageBoxIcon.Error,
                                          MessageBoxDefaultButton.Button2);
                            if (result == DialogResult.Yes)
                            {
                                FileInfo fi = new FileInfo(zipPath)
                                {
                                    CreationTime = dt,  // 作成日時
                                    LastWriteTime = dt, // 更新日時
                                    LastAccessTime = dt // アクセス日時
                                };

                                if (AddYMD) AddModifiedTimeToFileName(zipPath, dt);
                            }
                            else if (result == DialogResult.Cancel)
                            {
                                return "";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        return ex.Message;
                    }
                }
            }
            return "";
        }

        public static DateTime ZipGetLastModifiedFileTime(string zipPath, out string lastModifiedFileName)
        {
            lastModifiedFileName = "";
            DateTime dt = DateTime.MinValue;

            using (var fs = new FileStream(zipPath, FileMode.Open, FileAccess.Read))
            using (var zis = new ICSharpCode.SharpZipLib.Zip.ZipInputStream(fs))
            {
                while (true)
                {
                    var ze = zis.GetNextEntry();
                    if (ze == null) break;

                    if (ze.IsFile)
                    {
                        if (ze.DateTime > dt && (!Regex.IsMatch(ze.Name, @"thumbs\.db$", RegexOptions.IgnoreCase)))
                        {
                            dt = ze.DateTime;
                            lastModifiedFileName = ze.Name;
                        }
                    }
                }
            }

            //if (dt != DateTime.MinValue)
            //{
            //    Console.WriteLine("ZIPファイル名       : " + zipPath);
            //    Console.WriteLine("  ZIP内最新ファイル : " + lastModifiedFileName);
            //    Console.WriteLine("  その日付          : " + dt);
            //}

            return dt;
        }

        public static void ZipList(string zipPath)
        {
            var fs = new FileStream(zipPath, FileMode.Open, FileAccess.Read);
            var zis = new ICSharpCode.SharpZipLib.Zip.ZipInputStream(fs);

            while (true)
            {
                // ZipEntryを取得
                var ze = zis.GetNextEntry();
                if (ze == null) break;

                if (ze.IsFile)
                {
                    // ファイルのとき 
                    Console.WriteLine("名前 : {0}", ze.Name);
                    Console.WriteLine("サイズ : {0} bytes", ze.Size);
                    Console.WriteLine("格納サイズ : {0} bytes", ze.CompressedSize);
                    Console.WriteLine("圧縮方法 : {0}", ze.CompressionMethod);
                    Console.WriteLine("CRC : {0:X}", ze.Crc);
                    Console.WriteLine("日時 : {0}", ze.DateTime);
                    Console.WriteLine();
                }
                else if (ze.IsDirectory)
                {
                    // ディレクトリのとき 
                    Console.WriteLine("ディレクトリ名 : {0}", ze.Name);
                    Console.WriteLine("日時 : {0}", ze.DateTime);
                    Console.WriteLine();
                }
            }

            zis.Close();
            fs.Close();
        }

        /// <summary>
        /// ファイル名に日付を付与する（file.ext -> file_yymmdd.ext）
        /// </summary>
        /// <param name="path"></param>
        /// <param name="dt"></param>
        private static void AddModifiedTimeToFileName(string path, DateTime dt)
        {
            string ymd = dt.ToString("yyyyMMdd");
            string folder = Path.GetDirectoryName(path);
            string file = Path.GetFileName(path);
            int n = file.IndexOf(".");
            if (n < 0)
            {
                if (path.LastIndexOf($"_{ymd}") == -1)
                {
                    file = $"{file}_{ymd}";
                }
            }
            else
            {
                if (path.LastIndexOf($"_{ymd}{file.Substring(n)}") == -1)
                {
                    file = $"{file.Substring(0, n)}_{ymd}{file.Substring(n)}";
                }
            }

            File.Move(path, Path.Combine(folder, file));
        }
    }
}
