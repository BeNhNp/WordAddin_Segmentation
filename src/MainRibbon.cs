using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System.IO;

namespace WordAddIn_Segment
{
    public partial class MainRibbon
    {
        private int char_count;
        private int word_count;
        private int para_count;
        private Word.Document currentDocument;

        private void Segment()
        {
            currentDocument = Globals.ThisAddIn.Application.ActiveDocument;
            if (currentDocument.Paragraphs == null || currentDocument.Paragraphs.Count == 0)
                return;

            try
            {

                char_count = 0;
                word_count = 0;
                para_count = 0;

                int count;
                para_count = currentDocument.Paragraphs.Count;
                for (int i = para_count; i > 0; i--)
                {
                    String s = currentDocument.Paragraphs[i].Range.Text;
                    
                    byte[] bytes = System.Text.Encoding.Default.GetBytes(s);

                    count = ICTCLAS.NLPIR_GetParagraphProcessAWordCount(s);
                    result_t[] result = new result_t[count];//在客户端申请资源
                    ICTCLAS.NLPIR_ParagraphProcessAW(count, result);//获取结果存到客户的内存中

                    word_count += count;
                    char_count += s.Length;

                    String sn = System.Text.Encoding.Default.GetString(bytes, 0, result[0].length);
                    for (int j = 1; j < count; j++)
                    {
                        sn += ' ' + System.Text.Encoding.Default.GetString(bytes, result[j].start, result[j].length);

                    }
                    currentDocument.Paragraphs[i].Range.Text = sn;
                }
                
                //unknown bugs fix
                if (para_count + 1 == currentDocument.Paragraphs.Count)
                {
                    currentDocument.Paragraphs[currentDocument.Paragraphs.Count].Range.Delete();
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
                return;
            }

        }

        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            RegistryKey myKey = Registry.LocalMachine.OpenSubKey(@"Software\Demo", false);

            String dataPath = (String)myKey.GetValue("");
            String dir = Directory.GetCurrentDirectory();
            Directory.SetCurrentDirectory(dataPath);

            try
            {
                if (!ICTCLAS.NLPIR_Init("./", 0, ""))
                {
                    System.Windows.Forms.MessageBox.Show("Init ICTCLAS failed!");
                    return;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                throw;
            }
            Directory.SetCurrentDirectory(dir);
            currentDocument = null;
        }

        private void btn_Segment_Click(object sender, RibbonControlEventArgs e)
        {
            Segment();
        }

        private void btn_Statistics_Click(object sender, RibbonControlEventArgs e)
        {
            if(currentDocument==null)
            {
                currentDocument = Globals.ThisAddIn.Application.ActiveDocument;
                char_count = currentDocument.Characters.Count;
                word_count = currentDocument.Words.Count;
                para_count = currentDocument.Paragraphs.Count;
            }
            String msg = "";
            msg += "Characters: " + char_count.ToString();
            msg += "\nWords: " + word_count.ToString();
            msg += "\nParagraphs: " + para_count.ToString();
            System.Windows.Forms.MessageBox.Show(msg);
        }
    }
}
