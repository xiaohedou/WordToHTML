using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Checksums;
using System.Diagnostics;
using Microsoft.Office.Core;

namespace WordToHTML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        List<string> list = new List<string>();
        //获取指定目录下的所有文件以及文件夹
        private void GetAllFile(string strDir)
        {
            listView2.Items.Clear();
            string[] f = Directory.GetFileSystemEntries(strDir);
            string temp = "";
            for (int i = 0; i < f.Length; i++)
            {
                listView2.Items.Add(f[i].ToString());
                if (f[i].ToString().LastIndexOf(".") > 0)
                {
                    temp = f[i].ToString().Substring(0, f[i].ToString().LastIndexOf("."));
                    //if (!list.Contains(temp))
                        list.Add(f[i].ToString().Substring(0, f[i].ToString().LastIndexOf(".")));
                }
            }
        }
        //将word转换为html
        private void WordToHtmlFile(string WordFilePath,string strExtention)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application wApp = new Microsoft.Office.Interop.Word.Application();
                //指定原文件和目标文件 
                object docPath = WordFilePath;
                string htmlPath;
                FileInfo finfo = new FileInfo(WordFilePath);
                htmlPath = textBox2.Text.TrimEnd(new char[] { '\\' }) + "\\" + finfo.Name.Substring(0, finfo.Name.LastIndexOf(".")) + strExtention;
                object Target = htmlPath;
                //缺省参数 
                object Unknown = Type.Missing;
                //只读方式打开 
                object readOnly = true;
                //打开doc文件 
                Microsoft.Office.Interop.Word.Document document = wApp.Documents.Open(ref docPath, ref Unknown,
                ref readOnly, ref Unknown, ref Unknown,
                ref Unknown, ref Unknown, ref Unknown,
                ref Unknown, ref Unknown, ref Unknown,
                ref Unknown);
                // 指定格式
                object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFilteredHTML;
                switch (strExtention)
                {
                    case ".htm":
                    default:
                        format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFilteredHTML;
                        break;
                    case ".pdf":
                        format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;
                        break;
                    case ".rtf":
                        format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatRTF;
                        break;
                }
                object encoding = comboBox2.Text;
                switch (comboBox2.Text)
                {
                    case "GB2312":
                    default:
                        document.WebOptions.Encoding = MsoEncoding.msoEncodingSimplifiedChineseGBK;
                        break;
                    case "UTF8":
                        document.WebOptions.Encoding = MsoEncoding.msoEncodingUTF8;
                        break;
                }
                // 转换格式
                document.SaveAs(ref Target, ref format,
                ref Unknown, ref Unknown, ref Unknown,
                ref Unknown, ref Unknown, ref Unknown,
                ref Unknown, ref Unknown, ref Unknown);
                //document.SaveAs2(ref Target, ref format,
                //ref Unknown, ref Unknown, ref Unknown,
                //ref Unknown, ref Unknown, ref Unknown,
                //ref Unknown, ref Unknown, ref Unknown,
                //ref encoding, ref Unknown, ref Unknown,
                //ref Unknown, ref Unknown, ref Unknown);
                // 关闭文档和Word程序 
                document.Close(ref Unknown, ref Unknown, ref Unknown);
                wApp.Quit(ref Unknown, ref Unknown, ref Unknown);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        //批量转换
        private void BatchConvert(string docDir)
        {
            FileInfo finfo;
            //创建数组保存文件夹下的文件名 
            string[] docFiles = Directory.GetFiles(docDir);
            for (int i = 0; i < docFiles.Length; i++)
            {
                finfo = new FileInfo(docFiles[i]);
                if (finfo.Extension == ".doc" || finfo.Extension == ".docx")
                    WordToHtmlFile(docFiles[i],comboBox1.Text);
            }
        }

        void timer1_Tick(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "当前时间：" + DateTime.Now;//实时显示当前系统时间
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            if (folder.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folder.SelectedPath;
                string[] docFiles = Directory.GetFiles(textBox1.Text.Trim());
                FileInfo finfo;
                for (int i = 0; i < docFiles.Length; i++)
                {
                    finfo = new FileInfo(docFiles[i]);
                    if (finfo.Extension == ".doc" || finfo.Extension == ".docx")
                    {
                        listView1.Items.Add(docFiles[i].ToString());
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            folder.SelectedPath="d:/wordhtml";
            if (folder.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = folder.SelectedPath;//记录选择路径
                GetAllFile(textBox2.Text.Trim());
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count > 0 && textBox2.Text != "")
            {
                listView2.Items.Clear();//清空文件列表
                System.Threading.ThreadPool.QueueUserWorkItem(//使用线程池
                         (P_temp) =>
                         {
                             button4.Enabled = false;
                             BatchConvert(textBox1.Text.Trim());
                             GetAllFile(textBox2.Text.Trim());
                             MessageBox.Show("文档格式转换完成，快去使用吧！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                             button4.Enabled = true;
                         });
            }
            else
            {
                MessageBox.Show("请确认存在要转换的Word文档列表和转换后的文件存放路径！", "温馨提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void listView2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listView2.SelectedIndices.Count > 0)
                System.Diagnostics.Process.Start(listView2.SelectedItems[0].Text);
        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listView1.SelectedIndices.Count > 0)
                System.Diagnostics.Process.Start(listView1.SelectedItems[0].Text);
        }

        #region 压缩文件及文件夹
        ///// <summary>
        ///// 递归压缩文件夹方法
        ///// </summary>
        ///// <param name="FolderToZip"></param>
        ///// <param name="ZOPStream">压缩文件输出流对象</param>
        ///// <param name="ParentFolderName"></param>
        //private bool ZipFileDictory(string FolderToZip, ZipOutputStream ZOPStream, string ParentFolderName)
        //{
        //    bool res = true;
        //    string[] folders, filenames;
        //    ZipEntry entry = null;
        //    FileStream fs = null;
        //    Crc32 crc = new Crc32();
        //    try
        //    {
        //        //创建当前文件夹
        //        entry = new ZipEntry(Path.Combine(ParentFolderName, Path.GetFileName(FolderToZip) + "\\"));  //加上 “/” 才会当成是文件夹创建

        //        ZOPStream.PutNextEntry(entry);
        //        ZOPStream.Flush();
        //        //先压缩文件，再递归压缩文件夹 
        //        filenames = Directory.GetFiles(FolderToZip);
        //        foreach (string file in filenames)
        //        {
        //            //打开压缩文件
        //            fs = File.OpenRead(file);
        //            byte[] buffer = new byte[fs.Length];
        //            fs.Read(buffer, 0, buffer.Length);
        //            entry = new ZipEntry(Path.Combine(ParentFolderName, Path.GetFileName(FolderToZip) + "\\" + Path.GetFileName(file)));
        //            entry.DateTime = DateTime.Now;
        //            entry.Size = fs.Length;
        //            fs.Close();
        //            crc.Reset();
        //            crc.Update(buffer);
        //            entry.Crc = crc.Value;
        //            ZOPStream.PutNextEntry(entry);
        //            ZOPStream.Write(buffer, 0, buffer.Length);
        //        }
        //    }
        //    catch
        //    {
        //        res = false;
        //    }
        //    finally
        //    {
        //        if (fs != null)
        //        {
        //            fs.Close();
        //            fs = null;
        //        }
        //        if (entry != null)
        //        {
        //            entry = null;
        //        }
        //        GC.Collect();
        //    }
        //    folders = Directory.GetDirectories(FolderToZip);
        //    foreach (string folder in folders)
        //    {
        //        if (!ZipFileDictory(folder, ZOPStream, Path.Combine(ParentFolderName, Path.GetFileName(FolderToZip))))
        //        {
        //            return false;
        //        }
        //    }

        //    return res;
        //}

        ///// <summary>
        ///// 压缩文件夹
        ///// </summary>
        ///// <param name="FolderToZip">待压缩的文件夹</param>
        ///// <param name="ZipedFile">压缩后的文件名</param>
        ///// <returns></returns>
        //private bool ZipFileDictory(string FolderToZip, string ZipedFile)
        //{
        //    bool res;
        //    if (!Directory.Exists(FolderToZip))
        //    {
        //        return false;
        //    }
        //    ZipOutputStream ZOPStream = new ZipOutputStream(File.Create(ZipedFile));
        //    ZOPStream.SetLevel(6);
        //    res = ZipFileDictory(FolderToZip, ZOPStream, "");
        //    ZOPStream.Finish();
        //    ZOPStream.Close();
        //    return res;
        //}

        ///// <summary>
        ///// 压缩文件和文件夹
        ///// </summary>
        ///// <param name="FileToZip">待压缩的文件或文件夹</param>
        ///// <param name="ZipedFile">压缩后生成的压缩文件名，全路径格式</param>
        ///// <returns></returns>
        //public bool Zip(String FileToZip, String ZipedFile)
        //{
        //    if (Directory.Exists(FileToZip))
        //    {
        //        return ZipFileDictory(FileToZip, ZipedFile);
        //    }
        //    else
        //    {
        //        return false;
        //    }
        //}
        #endregion

        //批量压缩
        private void button3_Click(object sender, EventArgs e)
        {
            #region 拷贝到文件夹中批量压缩
            //DirectoryInfo dir = null;
            //System.Threading.ThreadPool.QueueUserWorkItem(//使用线程池，防止程序假死
            //             (P_temp) =>
            //             {
            //                 button3.Enabled = false;
            //                 try
            //                 {
            //                     if (listView2.Items.Count > 0)
            //                     {
            //                         for (int i = 0; i < list.Count; i++)
            //                         {
            //                             //if (i >= list.Count)
            //                             //    i -= list.Count;
            //                             dir = new DirectoryInfo(list[i]);
            //                             if (!dir.Exists)
            //                                 dir.Create();
            //                             ListViewItem li = listView2.Items.Cast<ListViewItem>().First(x => x.Text.Substring(0, x.Text.LastIndexOf(".")) == list[i]);
            //                             if (li != null)
            //                             {
            //                                 if (Directory.Exists(li.Text))
            //                                     Directory.Move(li.Text, (list[i] + "\\" + new DirectoryInfo(li.Text).Name));
            //                                 else if (File.Exists(li.Text))
            //                                     File.Move(li.Text, (list[i] + "\\" + new FileInfo(li.Text).Name));
            //                             }
            //                             listView2.Items.Remove(li);
            //                         }
            //                     }
            //                 }
            //                 catch { }
            //                 GetAllFile(textBox2.Text.Trim());
            //                 for (int i = 0; i < listView2.Items.Count; i++)
            //                 {
            //                     Zip(listView2.Items[i].Text.Replace('\\', '/'), (textBox2.Text.Replace('\\', '/').TrimEnd(new char[] { '/' }) + "/" + new DirectoryInfo(listView2.Items[i].Text).Name + ".zip").Replace('\\', '/'));
            //                 }
            //                 listView2.Items.Clear();
            //                 string[] f = Directory.GetFiles(textBox2.Text);
            //                 for (int i = 0; i < f.Length; i++)
            //                 {
            //                     listView2.Items.Add(f[i].ToString());
            //                 }
            //                 MessageBox.Show("压缩文件成功……", "恭喜你", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //                 button3.Enabled = true;
            //             });
            #endregion

            #region 调用批处理命令直接批量压缩
            Process proc = null;
            System.Threading.ThreadPool.QueueUserWorkItem(//使用线程池，防止程序假死
                         (P_temp) =>
                         {
                             proc = new Process();
                             proc.StartInfo.FileName = "zip.bat";
                             proc.StartInfo.Arguments = string.Format("10");
                             proc.StartInfo.UseShellExecute = false;
                             proc.StartInfo.CreateNoWindow = true;
                             proc.Start();
                             proc.WaitForExit();
                             MessageBox.Show("压缩文件成功……", "恭喜你", MessageBoxButtons.OK, MessageBoxIcon.Information);
                         });
            #endregion
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                textBox2.Text = "d:/wordhtml";
                DirectoryInfo dinfo = new DirectoryInfo("d:/wordhtml");
                if (!dinfo.Exists)
                    dinfo.Create();
            }
        }
    }
}