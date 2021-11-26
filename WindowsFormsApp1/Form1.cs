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
using System.IO;
using System.Collections;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public int countstring = 0;
        public Form1()
        {      
            InitializeComponent();
            dataGridView1.ContextMenuStrip = contextMenuStrip1;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabControl1.TabPages["tabPage1"];

            string[] row1 = new string[50];
            dataGridView1.Rows.Clear();
            dataGridView1.ColumnCount = 20;
            
            SQLiteConnection Connect = new SQLiteConnection(@"Data Source=D:/resume.db; Version=3;");
            SQLiteCommand command = Connect.CreateCommand();
            Connect.Open();
            command.CommandText = @"SELECT * FROM resume ";
            SQLiteDataReader reader = command.ExecuteReader();
            
            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    row1[i] = "";//чистка поля перед записью
                    row1[i] += reader.GetValue(i).ToString();
                    //richTextBox1.Text += reader.GetValue(i).ToString();
                    //richTextBox1.Text += reader.GetValue(0).ToString();
                    countstring++;
                }
                dataGridView1.Rows.Add(row1);
            }

        }
        Word._Application application;
        Word._Document document;
        Object missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;

        public string[] strmas = new string[100];

        private void openfile()
        {
            //создаем обьект приложения word
            application = new Word.Application();
            // создаем путь к файлу
            Object templatePathObj = "D:/resume.docx";

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

        //Считывание всего текста
        private void schitivanietekstaizdoka(string name)
        {
            string fileName = @name;
            string sdf;
            name=name.Replace('\'', '/');
            using (WordprocessingDocument myDocument = WordprocessingDocument.Open(fileName, true))
            {
                Body body = myDocument.MainDocumentPart.Document.Body;
                for (int g = 0; g < 19; g++)
                {
                    DocumentFormat.OpenXml.Wordprocessing.Paragraph firstParagraph = body.Elements<Paragraph>().ElementAt<Paragraph>(g);
                    DocumentFormat.OpenXml.OpenXmlElement firstChild = firstParagraph.FirstChild;
                    IEnumerable<Run> elementsAfter = firstChild.ElementsAfter().Where(c => c is Run).Cast<Run>();
                    foreach (DocumentFormat.OpenXml.Wordprocessing.Run runs in elementsAfter)
                    {
                        sdf = runs.InnerText.ToString();
                        strmas[g+1] = sdf;
                    }
                }
                myDocument.MainDocumentPart.Document.Save();
            }
        }

        private void ochistka()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            textBox16.Text = "";
            textBox17.Text = "";
            textBox18.Text = "";
            textBox19.Text = "";
        }

        private void soderzimoe_doca_dlya_pravki(string[] strmas1)
        {
            //string a = textBox20.Text;

            //schitivanietekstaizdoka("D:/resume.docx");

            textBox1.Text += strmas1[1];
            textBox2.Text += strmas1[2];
            textBox3.Text += strmas1[3];
            textBox4.Text += strmas1[4];
            textBox5.Text += strmas1[5];
            textBox6.Text += strmas1[6];
            textBox7.Text += strmas1[7];
            textBox8.Text += strmas1[8];
            textBox9.Text += strmas1[9];
            textBox10.Text += strmas1[10];
            textBox11.Text += strmas1[11];
            textBox12.Text += strmas1[12];
            textBox13.Text += strmas1[13];
            textBox14.Text += strmas1[14];
            textBox15.Text += strmas1[15];
            textBox16.Text += strmas1[16];
            textBox17.Text += strmas1[17];
            textBox18.Text += strmas1[18];
            textBox19.Text += strmas1[19];
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ochistka();
            soderzimoe_doca_dlya_pravki(strmas);

        }

        public string[] dirs1;
        private void schitivaniefailovvpapke()
        {
            //string[] dirs = Directory.GetFiles(@"D:\test");
            dirs1 = Directory.GetFiles(@"D:\test");
            foreach (string path in dirs1)
            {
                if (File.Exists(path))
                {
                    // This path is a file
                    richTextBox1.Text += path;
                    richTextBox1.Text += "\n";
                }
                else if (Directory.Exists(path))
                {
                    // This path is a directory
                    string[] fileEntries = Directory.GetFiles(path);

                }
            }
        }

        private void perenos_v_bd()
        {
            SQLiteConnection Connect = new SQLiteConnection(@"Data Source=D:/resume.db; Version=3;");
            SQLiteCommand command = Connect.CreateCommand();
            Connect.Open();
                command.CommandText = @"INSERT INTO resume ('Фамилия', 'Имя', 'Отчество', 'Адрес', 'Телефон', 'Цель', 'Образование', 'Диплом', 'Дата получения', 'Учебное заведение', 'Специализация', 'Дополнительная специализация', 'Курсовые работы по специальности', 'Навыки и умения', 'Управление', 'Опыт работы', 'Должность', 'Организация', 'Даты с – по') 
                VALUES ('" + textBox1.Text.ToString()+ "', '" + textBox2.Text.ToString() + "', '" + textBox3.Text.ToString() + "', '" + textBox4.Text.ToString() + "', '" + textBox5.Text.ToString() + "', '" + textBox6.Text.ToString() + "', '" + textBox7.Text.ToString() + "', '" + textBox8.Text.ToString() + "', '" + textBox9.Text.ToString() + "', '" + textBox10.Text.ToString() + "', '" + textBox11.Text.ToString() + "', '" + textBox12.Text.ToString() + "', '" + textBox13.Text.ToString() + "', '" + textBox14.Text.ToString() + "', '" + textBox15.Text.ToString() + "', '" + textBox16.Text.ToString() + "', '" + textBox17.Text.ToString() + "', '" + textBox18.Text.ToString() + "', '" + textBox19.Text.ToString() + "')";
                command.ExecuteReader();
            Connect.Close();
        }

        int stroka = 0;
        private void peredelka_bd()
        {
            SQLiteConnection Connect = new SQLiteConnection(@"Data Source=D:/resume.db; Version=3;");
            SQLiteCommand command;            
            Connect.Open();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET Фамилия = "+ "'" + textBox1.Text.ToString()+ "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET Имя = " + "'" + textBox2.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET Отчество = " + "'" + textBox3.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET Адрес = " + "'" + textBox4.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET Телефон = " + "'" + textBox5.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET Цель = " + "'" + textBox6.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET Образование = " + "'" + textBox7.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET Диплом = " + "'" + textBox8.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET 'Дата получения' = " + "'" + textBox9.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET 'Учебное заведение' = " + "'" + textBox10.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET Специализация = " + "'" + textBox11.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET 'Дополнительная специализация' = " + "'" + textBox12.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET 'Курсовые работы по специальности' = " + "'" + textBox13.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET 'Навыки и умения' = " + "'" + textBox14.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET Управление = " + "'" + textBox15.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET 'Опыт работы' = " + "'" + textBox16.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET Должность = " + "'" + textBox17.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET Организация = " + "'" + textBox18.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            command = Connect.CreateCommand();
            command.CommandText = @"UPDATE resume SET 'Даты с – по' = " + "'" + textBox19.Text.ToString() + "'" + "WHERE identificator = " + "'" + (stroka + 1).ToString() + "'";
            command.ExecuteReader();
            Connect.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            perenos_v_bd();
        }


        private void button5_Click_1(object sender, EventArgs e)
        {
            schitivaniefailovvpapke();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            schitivaniefailovvpapke();

            foreach (string path in dirs1)
            {

                string a = path;
                richTextBox1.Text += a;
                ochistka();
                schitivanietekstaizdoka(a);                
                soderzimoe_doca_dlya_pravki(strmas);
                perenos_v_bd();
            }
        }
        
        private void dataGridView1_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            
           stroka = e.RowIndex;
            //MessageBox.Show("sdfsdfsdf");
            ochistka();
            string[] str = new string[100];
            for(int i = 0; i < dataGridView1.CurrentRow.Cells.Count; i++)
            {
                str[i] = dataGridView1.CurrentRow.Cells[i].Value.ToString();
               
                //richTextBox1.Text += i;
            }
            soderzimoe_doca_dlya_pravki(str);
        }

        private void proverka_na_odinakovost()
        {
            SQLiteConnection Connect = new SQLiteConnection(@"Data Source=D:/resume.db; Version=3;");
            SQLiteCommand command = Connect.CreateCommand();
            Connect.Open();
            for(int i = 0; i < dataGridView1.RowCount; i++)
            command.CommandText = @"DELETE FROM resume [WHERE  LIKE " + а + "]";
            command.ExecuteReader();
            Connect.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            peredelka_bd();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {

        }
    }
    }

