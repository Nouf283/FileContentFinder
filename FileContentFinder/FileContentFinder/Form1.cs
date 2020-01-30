using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using TextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Code7248.word_reader;

namespace FileContentFinder
{
    public partial class Form1 : Form
    {

        string[] files;

        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (var fbd = new FolderBrowserDialog())
                {
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        files = Directory.GetFiles(fbd.SelectedPath);

                        if (files != null && files.Count() > 0)
                        {
                            textBox1.Text = string.Join(",", files);
                        }


                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(textBox2.Text) || string.IsNullOrEmpty(textBox2.Text))
                {
                    MessageBox.Show("Find Text cannot be empty");
                    return;
                }

                if (files == null || files.Count() <= 0)
                {
                    MessageBox.Show("Please select folder");
                    return;
                }
                foreach (string file in files)
                {
                    File.SetAttributes(file, FileAttributes.Normal);
                    string fileContent = File.ReadAllText(file);
                    if (fileContent.Contains(textBox2.Text))
                    {
                        richTextBox1.Text = richTextBox1.Text + "\n" + file;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = string.Empty;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private string GetTextFromPDF()
        {
            StringBuilder text = new StringBuilder();
            using (PdfReader reader = new PdfReader("D:\\RentReceiptFormat.pdf"))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }
            }

            return text.ToString();
        }

        //private string GetTextFromWord()
        //{
        //    StringBuilder text = new StringBuilder();
        //    Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
        //    object miss = System.Reflection.Missing.Value;
        //    object path = @"D:\Articles2.docx";
        //    object readOnly = true;
        //    Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

        //    for (int i = 0; i < docs.Paragraphs.Count; i++)
        //    {
        //        text.Append(" \r\n " + docs.Paragraphs[i + 1].Range.Text.ToString());
        //    }

        //    return text.ToString();
        //}

        private void readFileContent(string path)
        {
            TextExtractor extractor = new TextExtractor(path);
            string text = extractor.ExtractText();
            Console.WriteLine(text);
        }
    }
}
