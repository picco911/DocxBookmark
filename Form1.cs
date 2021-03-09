using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocxBookmark
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            insertAtBookmark();
        }

        public void insertAtBookmark()

        {
            var fis = new FileStream(@"e:\\temp\\bookmark.docx", FileMode.Open);
            var document = new XWPFDocument(fis);

            var paragraphList = new List<XWPFParagraph>();
            using (var paragraphs = document.GetParagraphsEnumerator())
            {
                while (paragraphs.MoveNext())
                {
                    paragraphList.Add(paragraphs.Current);
                }
            }


            foreach (var p in paragraphList)
            {
                //Here you have your paragraph; 
                var ctp = p.GetCTP();
                var changeItemList = new List<CT_Bookmark>();
                //iterate all paragraph items
                foreach (var item in ctp.Items)
                {
                    if (item.GetType().Name.Equals("CT_Bookmark"))
                    {
                        var bookmark = (CT_Bookmark)item;
                        if (bookmark.name.Equals("title"))
                        {
                            changeItemList.Add((CT_Bookmark)item);
                        }
                    }
                }

                //add new item and delete bookmark
                foreach (var item in changeItemList)
                {
                    XWPFRun run = p.CreateRun();
                    run.SetText("bookmark tsty");
                    ctp.Items.Remove(item);
                }
            }

            fis.Close();
            using (FileStream fs = new FileStream(@"e:\\temp\\bookmark1.docx", FileMode.CreateNew))
            {
                document.Write(fs);
            }

            document.Close();
        }

    }
}


