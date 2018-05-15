using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Saving;

namespace Word2jpg
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fl = new FolderBrowserDialog();
            if (fl.ShowDialog().Equals(DialogResult.OK))
            {
                textBox1.Text = fl.SelectedPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            /*string[] dirs = Directory.GetDirectories(textBox1.Text);
            string[] files = Directory.GetFiles(textBox1.Text);
            foreach (string file in files)
            {
                if (file.EndsWith(".doc") || file.EndsWith(".docx"))
                {
                    word2Jpg(file);
                }
            }

            foreach (string dir in dirs)
            {
                string[] dirFiles = Directory.GetFiles(dir);
                foreach (string dirFile in dirFiles)
                {
                    if (dirFile.EndsWith(".doc") || dirFile.EndsWith(".docx"))
                    {
                        word2Jpg(dirFile);
                    }
                }
            }*/
            Director(textBox1.Text);

            MessageBox.Show("导出完成");
        }

        private void word2Jpg(string docPath)
        {
            Document doc = new Document(docPath);
            textBox2.AppendText(string.Format("正在处理{0}\n", docPath));
            var rangeBookmark = doc.Range.Bookmarks["ConcordNumber"].Text.Trim();
            ImageSaveOptions iso = new ImageSaveOptions(SaveFormat.Jpeg);
            iso.Resolution = 300;
            iso.PrettyFormat = true;
            iso.UseAntiAliasing = true;
            for (int i = 0; i < 1; i++)
            {
                iso.PageIndex = i;
                //string path = string.Format("{0}/{1}CBHT{2}_{3}.jpg", textBox1.Text, rangeBookmark, doc.PageCount, i);
                string path = string.Format("{0}/{1}CBHT{2}_{3}.jpg", textBox1.Text, rangeBookmark, 2, i + 1);
                doc.Save(path, iso);
            }
        }

        private void Director(string dir)
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(dir);
            FileSystemInfo[] fsInfos = directoryInfo.GetFileSystemInfos();
            foreach (var info in fsInfos)
            {
                if (info is DirectoryInfo)
                {
                    Director(info.FullName);
                }
                else
                {
                    if (info.FullName.EndsWith(".doc") || info.FullName.EndsWith(".docx"))
                    {
                        word2Jpg(info.FullName);
                    }
                }
            }
        }
    }
}