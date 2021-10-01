using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            string[] row0 = { "11/22/1968", "29", "Revolution 9",
            "Beatles", "The Beatles [White Album]" };
            string[] row1 = new string[10];
            dataGridView1.ColumnCount = 5;
            dataGridView1.Columns[0].Name = "Release Date";
            dataGridView1.Columns[1].Name = "Track";
            dataGridView1.Columns[2].Name = "Title";
            dataGridView1.Columns[3].Name = "Artist";
            dataGridView1.Columns[4].Name = "Album";
            dataGridView1.Rows.Add(row0);


            SQLiteConnection Connect = new SQLiteConnection(@"Data Source=C:\Users\Kvantorium\Documents\sdfgh.db; Version=3;");
            SQLiteCommand command = Connect.CreateCommand();
            Connect.Open();
            command.CommandText =    @"SELECT * FROM sdhsg ";
            SQLiteDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {

                    richTextBox1.Text += reader.GetValue(i).ToString();
                    richTextBox1.Text += " ";
                    row1[i] = "";//чистка поля перед записью
                    row1[i] += reader.GetValue(i).ToString();


                }
                richTextBox1.Text += "\r\n";
                dataGridView1.Rows.Add(row1);
                

            }
           
        }
        Word._Application application;
        Word._Document document;
        Object missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;

        private void button1_Click(object sender, EventArgs e)
        {
            //создаем обьект приложения word
            application = new Word.Application();
            // создаем путь к файлу
            Object templatePathObj = "D:/prog.docx"; 

            // если вылетим не этом этапе, приложение останется открытым
            try
            {
                document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            }
            catch (Exception error)
            {
                document.Close(ref falseObj, ref missingObj, ref missingObj);
                application.Quit(ref missingObj, ref missingObj, ref missingObj);
                document = null;
                application = null;
                throw error;
            }
            application.Visible = true;


        }


        private void button2_Click(object sender, EventArgs e)
        {
            string fileName = @"D:/prog.docx";
            using (WordprocessingDocument myDocument = WordprocessingDocument.Open(fileName, true))
            {
                //Paragraph p =
                //    new Paragraph(
                //        new Run(
                //            new Text("This is some text in a run in a paragraph.")));
                //myDocument.MainDocumentPart.Document.Body.AppendChild(p);
                // Get the first paragraph.  
                //Paragraph p = myDocument.MainDocumentPart.Document.Body.Elements<Paragraph>().First();
                // If the paragraph has no ParagraphProperties object, create a new one.  
                //if (p.Elements<ParagraphProperties>().Count() == 0)
                //  p.PrependChild<ParagraphProperties>(new ParagraphProperties());
                // Get the ParagraphProperties element of the paragraph.  
                //ParagraphProperties pPr = p.Elements<ParagraphProperties>().First();
                // Set the value of ParagraphStyleId to "Heading3".  
                //pPr.ParagraphStyleId = new ParagraphStyleId() { Val = "Heading3" };

                Body body = myDocument.MainDocumentPart.Document.Body;
                Paragraph para = body.AppendChild(new Paragraph());
                //Run run = para.AppendChild(new Run());
                //run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"));
                //richTextBox1.Text += body.InnerText;

                //for(int h=0;h < body.Elements<Paragraph>().Count()-2;h++)
                //{
                //richTextBox1.Text += body.Elements<Paragraph>().Count<Paragraph>();


                RunProperties runProperties1 = new RunProperties();



                for (int g = 0; g < body.Elements<Paragraph>().Count<Paragraph>(); g++)
                {

                    DocumentFormat.OpenXml.Wordprocessing.Paragraph firstParagraph =
                            body.Elements<Paragraph>().ElementAt<Paragraph>(g);

                    DocumentFormat.OpenXml.OpenXmlElement firstChild = firstParagraph.FirstChild;


                    if (firstChild != null)
                    {
                        IEnumerable<Run> elementsAfter =
                    firstChild.ElementsAfter().Where(c => c is Run).Cast<Run>();



                        foreach (DocumentFormat.OpenXml.Wordprocessing.Run runs in elementsAfter)
                        {
                            richTextBox1.Text += runs.InnerText.ToString()+"\n";
                        }
                    }
                    else
                    {
                        richTextBox1.Text += "\n";

                    }
                }

                //Paragraph p = myDocument.MainDocumentPart.Document.Body.Descendants<Paragraph>().ElementAtOrDefault(0);
                // Call Save to generate an exception and show that access is read-only.
                myDocument.MainDocumentPart.Document.Save();


                  
                string text = "";


                //for (int i = 0; i < myDocument.Paragraphs.Count; i++)
                //{
                //    text += " \r\n " + myDocument.Paragraphs[i + 1].Range.Text;
                //}
                

                }
            MessageBox.Show("All done. Press a key.");
            
        }
    }
}
