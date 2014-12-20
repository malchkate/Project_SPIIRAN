using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


// Solution Explorer -> References -> [right click] -> Add Reference -> COM -> Microsoft Word... 
using Word = Microsoft.Office.Interop.Word; // использование псевдонимов
// Solution Explorer -> References -> [right click] -> Add Reference -> COM -> Microsoft Excel... 
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
namespace OurApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            tabControl1.TabPages[0].Name = "Добавление";
            tabControl1.TabPages[1].Name = "Извлечение";
        }

        private Word.Application wordapp;
        private Word.Document worddoc;
        private Excel.Application exapp;
        private Excel.Workbook exbook;
        public  List<Book> list_of_books = new List<Book>();
        string dbPath;
        int[] authorsLst  = new int[100];
        OleDbConnection  connection = new OleDbConnection();
        OleDbDataAdapter dataAdapter = new OleDbDataAdapter();
        DataSet ds;
        int numToAdd;   
        int? attribute_id_con = null;
        int? attribute_id_rep = null;
        private void BDButton_Click(object sender, EventArgs e)
        {
            Authors.Enabled = true;
            button2.Enabled = true;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "accdb files (*.accdb)|*.accdb";
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.ShowDialog();
            dbPath = openFileDialog1.FileName;
            pathLabel.Text = dbPath;

            try
            {
                //string queryString = "SELECT * FROM Общая LEFT JOIN Конференции ON Общая.Конференции = Конференции.Код;"; //JOIN Авторы
                string queryString = "SELECT * FROM [General]";
                string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath;
                connection.ConnectionString = conStr;
                dataAdapter = new OleDbDataAdapter(queryString, connection);
                ds = new DataSet();
                dataAdapter.Fill(ds, "General");
                // отображение на форме
                tableView.DataSource = ds.Tables["General"].DefaultView;
                Load();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error!:" + ex.Message + ex.StackTrace);
            }

        }

        private void Authors_Click(object sender, EventArgs e)
        {
            try
            {
                AuthorsForm form = new AuthorsForm(connection, dbPath, label5, ref authorsLst);
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error!:" + ex.Message + ex.StackTrace);
            }
        }

        // новая запись конференций
        private void button1_Click_2(object sender, EventArgs e)
        {
            try
            {
                string queryString = "SELECT * FROM [Attributes]";
                dataAdapter = new OleDbDataAdapter(queryString, connection);
                ds = new DataSet();
                dataAdapter.Fill(ds, "Attributes");
                int num = 0;
                foreach (DataRow row in ds.Tables["Attributes"].Rows)
                {
                    if (num < Convert.ToInt32(row[0]))
                        num = Convert.ToInt32(row[0]);
                }
                numToAdd = num + 1;
                attribute_id_con = numToAdd;
                var form = new ConForm(connection, dataAdapter, dbPath, label6, numToAdd);
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error!:" + ex.Message + ex.StackTrace);
            }
        }

        // Новая запись в Reports
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string queryString = "SELECT * FROM [Reports]";
                dataAdapter = new OleDbDataAdapter(queryString, connection);
                ds = new DataSet();
                dataAdapter.Fill(ds, "Reports");
                int num = 0;
                foreach (DataRow row in ds.Tables["Reports"].Rows)
                {
                    if (num < Convert.ToInt32(row[0]))
                        num = Convert.ToInt32(row[0]);
                }
                numToAdd = num + 1;
                attribute_id_rep = numToAdd;
                ReportForm form = new ReportForm(connection, dataAdapter, dbPath, numToAdd);
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error!:" + ex.Message + ex.StackTrace);
            }
        }

        // добавляем новую запись в General
        private void button2_Click(object sender, EventArgs e)
        {
            string str0 = category.Text;
            
            string str1 = name.Text;
            string str2 = year.Text;
            string str3 = pages.Text;
            string str4 = label.Text;
            string str5 = way.Text;
            string str6 = info.Text;
            try
            {
                connection.Open();

                string queryString = "SELECT * FROM [General]";
                dataAdapter = new OleDbDataAdapter(queryString, connection);
                ds = new DataSet();
                dataAdapter.Fill(ds, "General");
                OleDbCommand commandUpd = connection.CreateCommand();
                int num = 0;
                foreach (DataRow row in ds.Tables["General"].Rows)
                {
                    if (num < Convert.ToInt32(row[0]))
                        num = Convert.ToInt32(row[0]);
                }
                numToAdd = num + 1;
                string textQuery = "INSERT INTO [General] (id, category, pub_name, pub_year, pages_count, pub_label, url, info)" +
                                   " VALUES (" + numToAdd + ", '" + str0 + "', '" + str1 + "', '" + str2 + "', '" + str3 + "', '" + str4 + "', '" + str5 + "', '" + str6 + "')";
                if (attribute_id_rep != null)
                {
                    textQuery = "INSERT INTO [General] (id, category, pub_name, pub_year, pages_count, pub_label, url, info, attribute_id_rep)" +
                                  " VALUES (" + numToAdd + ", '" + str0 + "', '" + str1 + "', '" + str2 + "', '" + str3 + "', '" + str4 + "', '" + str5 + "', '" + str6 + "', " + attribute_id_rep + ")";
                } 
                else if (attribute_id_con != null)
                {
                    textQuery = "INSERT INTO [General] (id, category, pub_name, pub_year, pages_count, pub_label, url, info, attribute_id_con)" +
                                  " VALUES (" + numToAdd + ", '" + str0 + "', '" + str1 + "', '" + str2 + "', '" + str3 + "', '" + str4 + "', '" + str5 + "', '" + str6 + "', " + attribute_id_con + ")";
                }
                else if (attribute_id_con != null && attribute_id_rep != null)
                {
                    textQuery = "INSERT INTO [General] (id, category, pub_name, pub_year, pages_count, pub_label, url, info, attribute_id_rep, attribute_id_con)" +
                                  " VALUES (" + numToAdd + ", '" + str0 + "', '" + str1 + "', '" + str2 + "', '" + str3 + "', '" + str4 + "', '" + str5 + "', '" + str6 + "', " + attribute_id_rep + ", " + attribute_id_con + ")";
                }
                
                commandUpd.CommandText = textQuery;
                commandUpd.ExecuteNonQuery();

                queryString = "SELECT * FROM [AuthorVSpub]";
                dataAdapter = new OleDbDataAdapter(queryString, connection);
                ds = new DataSet();
                dataAdapter.Fill(ds, "AuthorVSpub");
                foreach (DataRow row in ds.Tables["AuthorVSpub"].Rows)
                {
                    num = Convert.ToInt32(row[0]);
                }
                for (int i = 0; authorsLst[i] != 0; i++)
                {
                    textQuery = "INSERT INTO [AuthorVSpub] (id, author, publication)" +
                                      " VALUES (" + (num + i + 1) + ", " + authorsLst[i] + ", " + numToAdd + ")";
                    commandUpd.CommandText = textQuery;
                    commandUpd.ExecuteNonQuery();
                }

                queryString = "SELECT * FROM [General]";
                dataAdapter = new OleDbDataAdapter(queryString, connection);
                ds = new DataSet();
                dataAdapter.Fill(ds, "General");
                tableView.DataSource = ds.Tables["General"].DefaultView;

                connection.Close();
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show("Error with connection:" + ex.Message);
            }
        }

        private void category_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((category.Text == "тезисы" || category.Text == "доклад на конференции") && String.Compare("*название базы данных*", pathLabel.Text) != 0)
            {
                button1.Enabled = true;
                button3.Enabled = false;
            }
            else if (category.Text == "отчет" && String.Compare("*название базы данных*", pathLabel.Text) != 0)
            {
                button3.Enabled = true;
                button1.Enabled = false;
            }
            else
            {
                button3.Enabled = false;
                button1.Enabled = false;
            }
        }
    
        private void Word_Table_MouseClick(object sender, MouseEventArgs e)
        {
            // создание нового файла со стандартным шаблоном 
            try
            {
                //Создаем объект Word - равносильно запуску Word 
                wordapp = new Word.Application();
                //Делаем его видимым 
                //wordapp.Visible = true;
                // открываем документ
                //worddoc = new Word.Document();
                worddoc = wordapp.Documents.Add();
                // textBox1.Enabled = false;


                worddoc = fill_Table(worddoc);
                //Делаем его видимым 
                wordapp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
                wordapp.Quit();
                worddoc = null;
                wordapp = null;

            }
        }

        string AUTHOR;
        private Word.Document fill_Table(Word.Document wodrdoc)
        {
            try
            {
               
                Get_Data_Table();
                // добавляем параграф
                worddoc.Paragraphs.Add();
                // выбираем первый параграф
                Word.Range wrange = worddoc.Paragraphs[1].Range;

                // добавляем текст
                wrange.Text = "Список \n научных работ \n " + AUTHOR;

                wrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                object objMiss = System.Reflection.Missing.Value;
                object objEndOfDocFlag = "\\endofdoc"; /* \endofdoc is a predefined bookmark */


                int n = list_of_books.Count + 2;

                Word.Table wordtable;
                Word.Range objWordRng = worddoc.Bookmarks.get_Item(ref objEndOfDocFlag).Range;
                wordtable = worddoc.Tables.Add(objWordRng, n, 6, ref objMiss, ref objMiss);

                int iRow, iCols;
                string strText;
                for (iRow = 1; iRow <= n; iRow++)
                    for (iCols = 1; iCols <= 6; iCols++)
                    {
                        //strText =  iRow + "c" + iCols;
                        //wordtable.Cell(iRow, iCols).Range.Text = strText;
                        wordtable.Cell(iRow, iCols).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                        wordtable.Cell(iRow, iCols).Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                        wordtable.Cell(iRow, iCols).Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                        wordtable.Cell(iRow, iCols).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;

                    }
                wordtable.Columns[1].PreferredWidth = 30;
                wordtable.Columns[2].PreferredWidth = 100;
                wordtable.Columns[6].PreferredWidth = 40;
                wordtable.Columns[5].PreferredWidth = 40;
                wordtable.Columns[3].PreferredWidth = 40;
                wordtable.Columns[4].PreferredWidth = 240;
                
                wordtable.Cell(1, 1).Range.Text = "1";
                wordtable.Cell(1, 2).Range.Text = "2";
                wordtable.Cell(1, 3).Range.Text = "3";
                wordtable.Cell(1, 4).Range.Text = "4";
                wordtable.Cell(1, 5).Range.Text = "5";
                wordtable.Cell(1, 6).Range.Text = "6";

                wordtable.Cell(2, 1).Range.Text = "№ \n п/п";
                wordtable.Cell(2, 2).Range.Text = "Наименование работы";
                wordtable.Cell(2, 3).Range.Text = "Форма работы";
                wordtable.Cell(2, 4).Range.Text = "Выходные данные";
                wordtable.Cell(2, 5).Range.Text = "Объем страниц ";
                wordtable.Cell(2, 6).Range.Text = "Соавторы";

                int i = 3;
                foreach (Book book in list_of_books)
                {

                    string[] elems = book.table_Word(AUTHOR);

                    wordtable.Cell(i, 1).Range.Text = "" + (i - 2).ToString();
                    for (iCols = 2; iCols <= 6; iCols++)
                    {
                        wordtable.Cell(i, iCols).Range.Text = elems[iCols - 2];
                    }

                    i++;
                }
                wordtable.Rows[1].Range.Font.Bold = 1;
                wordtable.Rows[1].Range.Font.Italic = 1;

                return worddoc;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
                wordapp.Quit();
                worddoc = null;
                wordapp = null;
                return null;

            }
        }

       

        private void Get_Data_Table()
        {
            try
            {
                if (Authors_List.SelectedItems.Count != 1)
                {
                    Console.WriteLine("ERROR.Only one Author could be selected!!!");
                }
                else
                {

                    string author = "";
                    foreach (String str in Authors_List.SelectedItems)
                    {
                        author = str;
                    }
                    AUTHOR = author;
                    string authorq = " WHERE [Authors].[author] = \"" + author + "\"";
                    if (Years_List.SelectedItems.Count > 0)
                    {

                        foreach (String year in Years_List.SelectedItems)
                        {
                            string yearsq = " AND [General].[pub_year] = \"" + year + "\"";
                            if (Conferentions_List.SelectedItems.Count > 0)
                            {
                                foreach (String conf in Conferentions_List.SelectedItems)
                                {
                                    string conferq = " AND Conferences.[conf_name] = \"" + conf + "\"";
                                    find_All(authorq, yearsq, conferq);     /////??????? and what about Reports and else???????
                                }
                            }
                            else
                            {
                                find_All(authorq, yearsq, "");
                            }
                        }
                    }
                    else
                    {
                        if (Conferentions_List.SelectedItems.Count > 0)
                        {
                            foreach (String conf in Conferentions_List.SelectedItems)
                            {
                                string conferq = " AND Conferences.[conf_name] = \"" + conf + "\"";
                                find_All(authorq, "", conferq);     /////??????? and what about Reports and else???????
                            }
                        }
                        else
                        {
                            find_All(authorq, "", "");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        public void find_All(string authorq, string yearsq, string conferq)
        {
            string querystring = "SELECT * FROM ((((Authors LEFT JOIN [AuthorVSpub] ON Authors.id = [AuthorVSpub].author) LEFT JOIN [General] ON [General].id = [AuthorVSpub].publication) LEFT JOIN Attributes ON [General].[attribute_id_con] = Attributes.id) LEFT JOIN Conferences ON Conferences.id = Attributes.[conference_id]) LEFT JOIN [Reports] ON [General].[attribute_id_rep] = [Reports].[id] ";
            querystring = querystring + authorq + yearsq + conferq;
            string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath;
            //connection.Open();
            //Console.WriteLine(connection.ConnectionString);
            //Console.WriteLine(connection.Database);
            dataAdapter = new OleDbDataAdapter(querystring, connection);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            System.Data.DataTable table = ds.Tables[0];
            foreach (DataRow row in table.Rows)
            {
                string iD = row.ItemArray[4].ToString();
                bool has_id = false;
                foreach (Book book in list_of_books)
                {
                    if (book.id.Equals(iD))
                    {
                        book.authors.Add(row.ItemArray[1].ToString());
                        // masha
                        has_id = true;
                    }
                }
                if (!has_id)
                {
                    string cat = row.ItemArray[6].ToString();
                    if (cat.Equals("тезисы") || cat.Equals("доклад на конференции"))
                    {
                        Tezis tezis = new Tezis();
                        List<string> co = new List<string>();
                        co = find_Coauth(iD);
                        tezis.fill(row, co);
                        list_of_books.Add(tezis);
                        list_of_books[0].authors.Distinct();
                    }
                    else if (cat.Equals("отчет"))
                    {
                        Otchet report = new Otchet();
                        List<string> co = new List<string>();
                        co = find_Coauth(iD);
                        report.fill(row, co);
                        list_of_books.Add(report);
                        list_of_books[0].authors.Distinct();
                    }
                    else if (cat.Equals("препринт"))
                    {

                        Preprint preprint = new Preprint();
                        List<string> co = new List<string>();
                        co = find_Coauth(iD);
                        preprint.fill(row, co);
                        list_of_books.Add(preprint);
                        list_of_books[0].authors.Distinct();
                    }

                }
            }
            connection.Close();
        }

        private void Get_Data_Word_List()
        {
            try
            {
                if (Authors_List.SelectedItems.Count > 0)
                {

                    foreach (String auth in Authors_List.SelectedItems)
                    {
                        string authorq = " WHERE ";
                        authorq = authorq + "[Authors].[author] = \"" + auth + "\"";


                        if (Categories_List.SelectedItems.Contains("тезисы"))
                        {
                            if (Years_List.SelectedItems.Count > 0)
                            {

                                foreach (String year in Years_List.SelectedItems)
                                {
                                    string yearsq = " AND [General].[pub_year] = \"" + year + "\"";
                                    if (Conferentions_List.SelectedItems.Count > 0)
                                    {
                                        foreach (String conf in Conferentions_List.SelectedItems)
                                        {
                                            string conferq = " AND Conferences.[conf_name] = \"" + conf + "\"";
                                            find_Tezis(authorq, yearsq, conferq);     /////??????? and what about Reports and else???????
                                        }
                                    }
                                    else
                                    {
                                        find_Tezis(authorq, yearsq, "");
                                    }
                                }
                            }
                            else
                            {
                                find_Tezis(authorq, "", "");
                            }
                        }

                        if (Categories_List.SelectedItems.Contains("доклад на конференции"))
                        {
                            if (Years_List.SelectedItems.Count > 0)
                            {

                                foreach (String year in Years_List.SelectedItems)
                                {
                                    string yearsq = " AND [General].[pub_year] = \"" + year + "\"";
                                    if (Conferentions_List.SelectedItems.Count > 0)
                                    {
                                        foreach (String conf in Conferentions_List.SelectedItems)
                                        {
                                            string conferq = " AND Conferences.[conf_name] = \"" + conf + "\"";
                                            find_Report_Conf(authorq, yearsq, conferq);     /////??????? and what about Reports and else???????
                                        }
                                    }
                                    else
                                    {
                                        find_Report_Conf(authorq, yearsq, "");
                                    }
                                }
                            }
                            else
                            {
                                find_Report_Conf(authorq, "", "");
                            }
                        }
                        if (Categories_List.SelectedItems.Contains("отчет"))
                        {
                            if (Years_List.SelectedItems.Count > 0)
                            {

                                foreach (String year in Years_List.SelectedItems)
                                {
                                    string yearsq = " AND [General].[pub_year] = \"" + year + "\"";
                                    find_Report(authorq, yearsq);
                                }
                            }
                            else
                            {
                                find_Report(authorq, "");
                            }


                        }
                        if (Categories_List.SelectedItems.Contains("препринт"))
                        {
                            if (Years_List.SelectedItems.Count > 0)
                            {

                                foreach (String year in Years_List.SelectedItems)
                                {
                                    string yearsq = " AND [General].[pub_year] = \"" + year + "\"";
                                    find_Preprint(authorq, yearsq);
                                }
                            }
                            else
                            {
                                find_Preprint(authorq, "");
                            }


                        }

                    }
                }
                else
                {
                    string authorq = " WHERE ";

                    if (Categories_List.SelectedItems.Contains("тезисы"))
                    {
                        if (Years_List.SelectedItems.Count > 0)
                        {

                            foreach (String year in Years_List.SelectedItems)
                            {
                                string yearsq = " AND [General].[pub_year] = \"" + year + "\"";
                                if (Conferentions_List.SelectedItems.Count > 0)
                                {
                                    foreach (String conf in Conferentions_List.SelectedItems)
                                    {
                                        string conferq = " AND Conferences.[conf_name] = \"" + conf + "\"";
                                        find_Tezis(authorq, yearsq, conferq);     /////??????? and what about Reports and else???????
                                    }
                                }
                                else
                                {
                                    find_Tezis(authorq, yearsq, "");
                                }
                            }
                        }
                        else
                        {
                            find_Tezis(authorq, "", "");
                        }


                    }
                    if (Categories_List.SelectedItems.Contains("доклад на конференции"))
                    {
                        if (Years_List.SelectedItems.Count > 0)
                        {

                            foreach (String year in Years_List.SelectedItems)
                            {
                                string yearsq = " AND [General].[pub_year] = \"" + year + "\"";
                                if (Conferentions_List.SelectedItems.Count > 0)
                                {
                                    foreach (String conf in Conferentions_List.SelectedItems)
                                    {
                                        string conferq = " AND Conferences.[conf_name] = \"" + conf + "\"";
                                        find_Report_Conf(authorq, yearsq, conferq);     /////??????? and what about Reports and else???????
                                    }
                                }
                                else
                                {
                                    find_Report_Conf(authorq, yearsq, "");
                                }
                            }
                        }
                        else
                        {
                            find_Report_Conf(authorq, "", "");
                        }
                    }
                    if (Categories_List.SelectedItems.Contains("отчет"))
                    {
                        if (Years_List.SelectedItems.Count > 0)
                        {

                            foreach (String year in Years_List.SelectedItems)
                            {
                                string yearsq = " AND [General].[pub_year] = \"" + year + "\"";
                                find_Report(authorq, yearsq);
                            }
                        }
                        else
                        {
                            find_Report(authorq, "");
                        }


                    }
                    if (Categories_List.SelectedItems.Contains("препринт"))
                    {
                        if (Years_List.SelectedItems.Count > 0)
                        {

                            foreach (String year in Years_List.SelectedItems)
                            {
                                string yearsq = " AND [General].[pub_year] = \"" + year + "\"";
                                find_Preprint(authorq, yearsq);
                            }
                        }
                        else
                        {
                            find_Preprint(authorq, "");
                        }


                    }

                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        private void Word_List_MouseClick(object sender, MouseEventArgs e)
        {
            // создание нового файла со стандартным шаблоном 
            try
            {
                Get_Data_Word_List();

                //Создаем объект Word - равносильно запуску Word 
                wordapp = new Word.Application();
                //Делаем его видимым 
                //wordapp.Visible = true;
                // открываем документ
                worddoc = wordapp.Documents.Add();
                // textBox1.Enabled = false;

                // добавляем параграф
                worddoc.Paragraphs.Add();
                // выбираем первый параграф
                Word.Range wrange = worddoc.Paragraphs[1].Range;
                wrange.Text = " Список \n опубликованных работ  \n ";
                //wrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                Word.Range wwrange = worddoc.Paragraphs[2].Range;
                wwrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                
                // добавляем текст
                foreach (string year in Years_List.Items)
                {
                    bool has_year = false;
                    foreach (Book elem in list_of_books)
                    {
                        if (elem.year.Equals(year))
                        {
                            has_year = true;
                            break;
                        }
                    }
                    if (has_year)
                    {
                        int i = 1;
                        Console.WriteLine("Year:" + year);
                        wwrange.Text += "\n " + year + "\n";
                        foreach (Book elem in list_of_books)
                        {
                            if (elem.year.Equals(year))
                            {
                                wwrange.Text += i + ") " + elem.list_Word() + "\n";
                                i++;
                            }
                        }
                    }
                }
                object objMiss = System.Reflection.Missing.Value;
                object objEndOfDocFlag = "\\endofdoc"; /* \endofdoc is a predefined bookmark */


                //Делаем его видимым 
                wordapp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
                wordapp.Quit();
                worddoc = null;
                wordapp = null;

            }
        }


        public void find_Tezis(string authorq, string yearsq, string conferq)
        {
            string categq = " AND [General].category = \"тезисы\"";
            string querystring = "SELECT * FROM ((((Authors LEFT JOIN [AuthorVSpub] ON Authors.id = [AuthorVSpub].author) LEFT JOIN [General] ON [General].id = [AuthorVSpub].publication) LEFT JOIN Attributes ON [General].[attribute_id_con] = Attributes.id) LEFT JOIN Conferences ON Conferences.id = Attributes.[conference_id]) LEFT JOIN [Reports] ON [General].[attribute_id_rep] = [Reports].[id] ";
            querystring = querystring + authorq + categq + yearsq + conferq;
            Console.WriteLine(querystring);
            string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath;
            //connection.Open();
            //Console.WriteLine(connection.ConnectionString);
            //Console.WriteLine(connection.Database);
            dataAdapter = new OleDbDataAdapter(querystring, connection);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            System.Data.DataTable table = ds.Tables[0];
            foreach (DataRow row in table.Rows)
            {
                string iD = row.ItemArray[4].ToString();
                bool has_id = false;
                //foreach (Book book in list_of_books)
                //{
                //    Console.WriteLine(book.id);
             
                //}
                
                foreach (Book book in list_of_books)
                {

                    if (book.id.Equals(iD))
                    {
                        book.authors.Add(row.ItemArray[1].ToString());
                        has_id = true;
                    }
                }
                if (!has_id)
                {
                    Tezis tezis = new Tezis();
                    List<string> co = new List<string>();
                    co = find_Coauth(iD);
                    tezis.fill(row, co);
                    list_of_books.Add(tezis);
                }

            }
            connection.Close();
        }

        public void find_Report_Conf(string authorq, string yearsq, string conferq)
        {
            string categq = " AND [General].category = \"доклад на конференции\"";
            string querystring = "SELECT * FROM ((((Authors LEFT JOIN [AuthorVSpub] ON Authors.id = [AuthorVSpub].author) LEFT JOIN [General] ON [General].id = [AuthorVSpub].publication) LEFT JOIN Attributes ON [General].[attribute_id_con] = Attributes.id) LEFT JOIN Conferences ON Conferences.id = Attributes.[conference_id]) LEFT JOIN [Reports] ON [General].[attribute_id_rep] = [Reports].[id] ";
            querystring = querystring + authorq + categq + yearsq + conferq;
            Console.WriteLine(querystring);
            string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath;
            //connection.Open();
            //Console.WriteLine(connection.ConnectionString);
            //Console.WriteLine(connection.Database);
            dataAdapter = new OleDbDataAdapter(querystring, connection);
            DataSet ds = new DataSet();
            dataAdapter.Fill(ds);
            System.Data.DataTable table = ds.Tables[0];
            foreach (DataRow row in table.Rows)
            {
                string iD = row.ItemArray[4].ToString();
                bool has_id = false;
                foreach (Book book in list_of_books)
                {
                    if (book.id.Equals(iD))
                    {
                        book.authors.Add(row.ItemArray[1].ToString());
                    }
                }
                if (!has_id)
                {
                    Tezis tezis = new Tezis();
                    List<string> co = new List<string>();
                    co = find_Coauth(iD);
                    tezis.fill(row, co);

                    list_of_books.Add(tezis);
                }


            }
            connection.Close();
        }

        public void find_Report(string authorq, string yearsq)
        {
            try
            {
                string categq = " AND [General].[category] = \"отчет\"";
                string querystring = "SELECT * FROM ((((Authors LEFT JOIN [AuthorVSpub] ON Authors.id = [AuthorVSpub].author) LEFT JOIN [General] ON [General].id = [AuthorVSpub].publication) LEFT JOIN Attributes ON [General].[attribute_id_con] = Attributes.id) LEFT JOIN Conferences ON Conferences.id = Attributes.[conference_id]) LEFT JOIN [Reports] ON [General].[attribute_id_rep] = [Reports].[id] ";
                querystring = querystring + authorq + categq + yearsq;
                string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath;
                //connection.Open();
                //Console.WriteLine(connection.ConnectionString);
                //Console.WriteLine(connection.Database);
                dataAdapter = new OleDbDataAdapter(querystring, connection);
                DataSet ds = new DataSet();
                dataAdapter.Fill(ds);
                System.Data.DataTable table = ds.Tables[0];
                foreach (DataRow row in table.Rows)
                {
                    string iD = row.ItemArray[4].ToString();
                bool has_id = false;
                foreach (Book book in list_of_books)
                {
                    if (book.id.Equals(iD))
                    {
                        book.authors.Add(row.ItemArray[1].ToString());
                    }
                }
                if (!has_id)
                {
                    Otchet report = new Otchet();
                    List<string> co = new List<string>();
                    co = find_Coauth(iD);
                    report.fill(row, co);
                    list_of_books.Add(report);
                }

                }
                connection.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("findreport " + ex.Message + ex.StackTrace);
            }
        }


        public List<string> find_Coauth(string id )
        {
            List<string> list_co = new List<string>();
            string querystring = "SELECT * FROM  [AuthorVSpub] LEFT JOIN Authors ON Authors.id = [AuthorVSpub].author WHERE [AuthorVSpub].publication =" + id;
            string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath;
                connection.Open();
                //Console.WriteLine(connection.ConnectionString);
                //Console.WriteLine(connection.Database);
               dataAdapter = new OleDbDataAdapter(querystring, connection);
            //masha
            //OleDbCommand command = connection.CreateCommand();
            // текст запроса
            //command.CommandText = querystring;
            // выполнение запроса
            //dataAdapter.SelectCommand = command;

                DataSet ds = new DataSet();
                dataAdapter.Fill(ds);
                System.Data.DataTable table = ds.Tables[0];
                foreach (DataRow row in table.Rows)
                {
                    list_co.Add(row.ItemArray[4].ToString());

                }
                connection.Close();
            return list_co;
            
        }


   
        public void find_Preprint(string authorq, string yearsq)
        {
            try
            {
                string categq = " AND [General].[category] =  \"препринт\"";
                string querystring = "SELECT * FROM ((((Authors LEFT JOIN [AuthorVSpub] ON Authors.id = [AuthorVSpub].author) LEFT JOIN [General] ON [General].id = [AuthorVSpub].publication) LEFT JOIN Attributes ON [General].[attribute_id_con] = Attributes.id) LEFT JOIN Conferences ON Conferences.id = Attributes.[conference_id]) LEFT JOIN [Reports] ON [General].[attribute_id_rep] = [Reports].[id] ";
                querystring = querystring + authorq + categq + yearsq;
                
                string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath;
                //connection.Open();
                //Console.WriteLine(connection.ConnectionString);
                //Console.WriteLine(connection.Database);
                dataAdapter = new OleDbDataAdapter(querystring, connection);
                DataSet ds = new DataSet();
                dataAdapter.Fill(ds);
                System.Data.DataTable table = ds.Tables[0];
                foreach (DataRow row in table.Rows)
                {
                    string iD = row.ItemArray[4].ToString();
                    bool has_id = false;
                    foreach (Book book in list_of_books)
                    {
                        if (book.id.Equals(iD))
                        {
                            book.authors.Add(row.ItemArray[1].ToString());
                        }
                    }
                    if (!has_id)
                    {
                        Preprint preprint = new Preprint();
                        List<string> co = new List<string>();
                          co =  find_Coauth(iD);
                        preprint.fill(row, co);
                        list_of_books.Add(preprint);

                    }
                }
                connection.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("findPreprint " + ex.Message + ex.StackTrace);
            }
        }

        private void Get_Excel_Data()
        {
            try
            {
                if (Authors_List.SelectedItems.Count > 0)
                {

                    foreach (String auth in Authors_List.SelectedItems)
                    {
                        string authorq = " WHERE ";
                        authorq = authorq + "[Authors].[author] = \"" + auth + "\"";

                        if (Years_List.SelectedItems.Count > 0)
                        {

                            foreach (String year in Years_List.SelectedItems)
                            {
                                string yearsq = " AND [General].[pub_year] = \"" + year + "\"";
                                if (Conferentions_List.SelectedItems.Count > 0)
                                {
                                    foreach (String conf in Conferentions_List.SelectedItems)
                                    {
                                        string conferq = " AND Conferences.[conf_name] = \"" + conf + "\"";
                                        find_All(authorq, yearsq, conferq);     /////??????? and what about Reports and else???????
                                    }
                                }
                                else
                                {
                                    find_All(authorq, yearsq, "");
                                }
                            }
                        }
                        else
                        {
                            if (Conferentions_List.SelectedItems.Count > 0)
                            {
                                foreach (String conf in Conferentions_List.SelectedItems)
                                {
                                    string conferq = " AND Conferences.[conf_name] = \"" + conf + "\"";
                                    find_All(authorq, "", conferq);     /////??????? and what about Reports and else???????
                                }
                            }
                            else
                            {
                                find_All(authorq, "", "");
                            }
                        }
                    }
                }
                else
                {
                    string authorq = " WHERE ";
                    if (Years_List.SelectedItems.Count > 0)
                    {

                        foreach (String year in Years_List.SelectedItems)
                        {
                            string yearsq = " AND [General].[pub_year] = \"" + year + "\"";
                            if (Conferentions_List.SelectedItems.Count > 0)
                            {
                                foreach (String conf in Conferentions_List.SelectedItems)
                                {
                                    string conferq = " AND Conferences.[conf_name] = \"" + conf + "\"";
                                    find_All(authorq, yearsq, conferq);     /////??????? and what about Reports and else???????
                                }
                            }
                            else
                            {
                                find_All(authorq, yearsq, "");
                            }
                        }
                    }
                    else
                    {
                        if (Conferentions_List.SelectedItems.Count > 0)
                        {
                            foreach (String conf in Conferentions_List.SelectedItems)
                            {
                                string conferq = " AND Conferences.[conf_name] = \"" + conf + "\"";
                                find_All(authorq, "", conferq);     /////??????? and what about Reports and else???????
                            }
                        }
                        else
                        {
                            find_All(authorq, "", "");
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        private void Excel_MouseClick(object sender, MouseEventArgs e)
        {
            //Excel.Application app;
            try
            {
                //Создаем объект Excel - равносильно запуску Excel 
                exapp = new Excel.Application();

                Get_Excel_Data();
                // создание книги с 3-мя листами
                exapp.SheetsInNewWorkbook = 4;
                exbook = exapp.Workbooks.Add();
                exbook = fill_Excel(exbook);
                //Делаем его видимым 
                exapp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                exapp.Quit();
                exbook = null;
                exapp = null;

            }

        }

        private Excel.Workbook fill_Excel(Excel.Workbook exbook)
        {
            try
            {
                //Получаем ссылку на лист 1 
                Excel.Worksheet excelws = exbook.Worksheets[1] ;
                excelws.Name = "Тезисы";
                int i = 1; // report
                Excel.Range excelcells = excelws.Cells[i, 1];
                excelcells.Value = "Авторы";
                excelcells = excelws.Cells[i, 2];
                excelcells.Value = "Название";
                excelcells = excelws.Cells[i, 3];
                excelcells.Value = "Год";
                excelcells = excelws.Cells[i, 4];
                excelcells.Value = "Общее количество стр";
                excelcells = excelws.Cells[i, 5];
                excelcells.Value = "Метка";
                excelcells = excelws.Cells[i, 6];
                excelcells.Value = "Ссылка";
                excelcells = excelws.Cells[i, 7];
                excelcells.Value = "Доп. информация";
                excelcells = excelws.Cells[i, 8];
                excelcells.Value = "Город";
                excelcells = excelws.Cells[i, 9];
                excelcells.Value = "Издательство";
                excelcells = excelws.Cells[i, 10];
                excelcells.Value = "Первая страница";
                excelcells = excelws.Cells[i, 11];
                excelcells.Value = "Последняя страница";
                excelcells = excelws.Cells[i, 12];
                excelcells.Value = "Том";
                excelcells = excelws.Cells[i, 13];
                excelcells.Value = "Название сборника";
                excelcells = excelws.Cells[i, 14];
                excelcells.Value = "Сокр. имя конференции";
                excelcells = excelws.Cells[i, 15];
                excelcells.Value = "Место конференции";
                excelcells = excelws.Cells[i, 16];
                excelcells.Value = "Дата конференции";

                //2
                excelws = exbook.Worksheets[2];
                excelws.Name = "Доклады на конференции";
                excelcells = excelws.Cells[i, 1];
                excelcells.Value = "Авторы";
                excelcells = excelws.Cells[i, 2];
                excelcells.Value = "Название";
                excelcells = excelws.Cells[i, 3];
                excelcells.Value = "Год";
                excelcells = excelws.Cells[i, 4];
                excelcells.Value = "Общее количество стр";
                excelcells = excelws.Cells[i, 5];
                excelcells.Value = "Метка";
                excelcells = excelws.Cells[i, 6];
                excelcells.Value = "Ссылка";
                excelcells = excelws.Cells[i, 7];
                excelcells.Value = "Доп. информация";
                excelcells = excelws.Cells[i, 8];
                excelcells.Value = "Город";
                excelcells = excelws.Cells[i, 9];
                excelcells.Value = "Издательство";
                excelcells = excelws.Cells[i, 10];
                excelcells.Value = "Первая страница";
                excelcells = excelws.Cells[i, 11];
                excelcells.Value = "Последняя страница";
                excelcells = excelws.Cells[i, 12];
                excelcells.Value = "Том";
                excelcells = excelws.Cells[i, 13];
                excelcells.Value = "Название сборника";
                excelcells = excelws.Cells[i, 14];
                excelcells.Value = "Сокр. имя конференции";
                excelcells = excelws.Cells[i, 15];
                excelcells.Value = "Место конференции";
                excelcells = excelws.Cells[i, 16];
                excelcells.Value = "Дата конференции";

                //3
                excelws = exbook.Worksheets[3];
                excelws.Name = "Отчеты";
                excelcells = excelws.Cells[i, 1];
                excelcells.Value = "Авторы";
                excelcells = excelws.Cells[i, 2];
                excelcells.Value = "Название";
                excelcells = excelws.Cells[i, 3];
                excelcells.Value = "Год";
                excelcells = excelws.Cells[i, 4];
                excelcells.Value = "Общее количество стр";
                excelcells = excelws.Cells[i, 5];
                excelcells.Value = "Метка";
                excelcells = excelws.Cells[i, 6];
                excelcells.Value = "Ссылка";
                excelcells = excelws.Cells[i, 7];
                excelcells.Value = "Доп. информация";
                excelcells = excelws.Cells[i, 8];
                excelcells.Value = "Город";
                excelcells = excelws.Cells[i, 9];
                excelcells.Value = "Издательство";
                excelcells = excelws.Cells[i, 10];
                excelcells.Value = "Инф. номер";
                excelcells = excelws.Cells[i, 11];
                excelcells.Value = "Рег. номер";
                excelcells = excelws.Cells[i, 12];
                excelcells.Value = "Дата рег.";
                excelcells = excelws.Cells[i, 13];
                excelcells.Value = "Тема";
                excelcells = excelws.Cells[i, 14];
                excelcells.Value = "Статус отчета";

                //4th page
                excelws = exbook.Worksheets[4];
                excelws.Name = "Препринты";
                excelcells = excelws.Cells[i, 1];
                excelcells.Value = "Авторы";
                excelcells = excelws.Cells[i, 2];
                excelcells.Value = "Название";
                excelcells = excelws.Cells[i, 3];
                excelcells.Value = "Год";
                excelcells = excelws.Cells[i, 4];
                excelcells.Value = "Общее количество стр";
                excelcells = excelws.Cells[i, 5];
                excelcells.Value = "Метка";
                excelcells = excelws.Cells[i, 6];
                excelcells.Value = "Ссылка";
                excelcells = excelws.Cells[i, 7];
                excelcells.Value = "Доп. информация";
                excelcells = excelws.Cells[i, 8];
                foreach (Book book in list_of_books)
                {
                    Console.WriteLine("BOOKSS" + book.category);
                    Console.WriteLine("BOOKSS" + (book.category.Equals("доклад на конференции")));
                }

                int tez = 2;
                int d_con = 2;
                int rep = 2;
                int prepr = 2;
                foreach (Book book in list_of_books)
                {
                    if (book.category.Equals("тезисы"))
                    {
                        excelws = exbook.Worksheets[1];
                        excelws = book.Excel(excelws, tez);
                        tez++;
                    }
                    else if (book.category.Equals("доклад на конференции"))
                    {
                        excelws = exbook.Worksheets[2];
                        excelws = book.Excel(excelws, d_con);
                        d_con++;
                    }
                    else if (book.category.Equals("отчет"))
                    {
                        excelws = exbook.Worksheets[3];
                        excelws = book.Excel(excelws, rep);
                        rep++;
                    }
                    else if (book.category.Equals("препринт"))
                    {
                        excelws = exbook.Worksheets[4];
                        excelws = book.Excel(excelws, prepr);
                        prepr++;
                    }
                }
                Console.WriteLine("NUMBERS" + tez + " " + d_con + "  " + rep + " " + prepr);

                return exbook;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                exapp.Quit();
                exbook = null;
                exapp = null;
                return null;
            }
        }


        private void Load()
        {
            try
            {



                //add Authors
                string querystring = "SELECT * FROM Authors";
                string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath;
                connection.ConnectionString = conStr;
                dataAdapter = new OleDbDataAdapter(querystring, connection);
                DataSet ds = new DataSet();
                dataAdapter.Fill(ds, "Authors");
                System.Data.DataTable table = ds.Tables["Authors"];
                foreach (DataRow row in table.Rows)
                {
                    Authors_List.Items.Add(row[1]);
                }


                //add Years
                querystring = "SELECT * FROM [General]";
                List<int> years = new List<int>();

                dataAdapter = new OleDbDataAdapter(querystring, connection);
                ds = new DataSet();
                dataAdapter.Fill(ds, "[General]");
                table = ds.Tables["[General]"];
                foreach (DataRow row in table.Rows)
                {
                    int temp = Convert.ToInt32(row.ItemArray[3]);
                    if (!years.Contains(temp))
                    {
                        years.Add(temp);
                    }
                }
                foreach (int year in years)
                {
                    Years_List.Items.Add(year.ToString());
                }

                //add Categories

                Categories_List.Items.Add("тезисы");
                Categories_List.Items.Add("доклад на конференции");
                Categories_List.Items.Add("отчет");
                Categories_List.Items.Add("препринт");

                //add Conferentions
                querystring = "SELECT * FROM Conferences";
                dataAdapter = new OleDbDataAdapter(querystring, connection);
                ds = new DataSet();
                dataAdapter.Fill(ds, "Conferences");
                table = ds.Tables["Conferences"];
                foreach (DataRow row in table.Rows)
                {
                    Conferentions_List.Items.Add(row[1]);
                }




            }
            catch (Exception ex)
            {
                MessageBox.Show("Error!:" + ex.Message + ex.StackTrace);
            }
        }

        private void Latex_button_MouseClick(object sender, MouseEventArgs e)
        {
            Get_Data_Word_List();
            TexThesis tex = new TexThesis(@"C:\Users\Elfimova\Documents\учеба\ООиИАД\project\OurApplication");
            tex.writeBibliographyThesis(list_of_books);
        }
    }
    


    abstract public class Book
    {
        public List<string> authors;
        public string name;
        public string year;
        public string link;
        public string extra_inf;
        public string city;
        public string category;
        public string publish_house;
        public string id;
        public Book()
        {
            this.authors = new List<string>();
        }
        abstract public string list_Word();
        abstract public Excel.Worksheet Excel(Excel.Worksheet excelws, int i);
        abstract public string[] table_Word(string name_a);
        abstract public void fill(DataRow row, List<string> co);
    }

    public class Tezis : Book
    {

        public string full_pages_number;
        public string label;
        public string first_page;
        public string last_page;
        public string part; //tom
        public string name_of_collection;
        public string short_name_conf;
        public string place_conf;
        public string date_conf;


        public override string list_Word()
        {
            string result = "";
            foreach (string author in this.authors)
            {
                result = result + author + ", ";// later make so, that it wouldnt print "," after last author;
            }
            int len = result.Length;
            result = result.Substring(0, len - 2);
            result = result + " " + this.name + " //";
            result = result + this.name_of_collection + ". ";
            if (this.part != null)
            {
                result = result + "T." + this.part + ". ";
            }
            if ((this.place_conf != null) && (this.date_conf != null))
            {
                result = result + "(" + this.place_conf + ", " + this.date_conf + "). ";
            }
            result = result + this.city + ": " + this.publish_house + ", " + this.year + ". C. " + this.first_page;
            if (this.last_page != null)
            {
                result = result + "-" + this.last_page + ". ";
            }
            else
            {
                result = result + ". ";
            }
            if (this.link != null)
            {
                result = result + "URL: " + this.link + ". ";
            }
            if (this.extra_inf != null)
            {
                result = result + "(" + this.extra_inf + ") ";
            }
            return result;
        }
        public override string[] table_Word(string name_a)
        {
            string[] elems = new string[5];
            elems[0] = this.name + "(" + this.category + ")";
            elems[1] = this.label;
            Console.WriteLine("NAAAME" + this.name);
            string result = this.name_of_collection + ". ";
            if (this.part != null)
            {
                result = result + "T." + this.part + ". ";
            }
            if ((this.place_conf != null) && (this.date_conf != null))
            {
                result = result + "(" + this.place_conf + ", " + this.date_conf + "). ";
            }
            result = result + this.city + ": " + this.publish_house + ", " + this.year + ". C. ";

            if (this.link != null)
            {
                result = result + "URL: " + this.link + ". ";
            }
            if (this.extra_inf != null)
            {
                result = result + "(" + this.extra_inf + ") ";
            }
            elems[2] = result;
            elems[3] = this.full_pages_number;
            elems[4] = "";
            foreach (string auth in this.authors)
            {
                if (!auth.Equals(name_a))
                {
                    elems[4] = auth + ", ";
                }
            }

            
            return elems;
        }
        public override void fill(DataRow row, List<string> co)
        {
            //masha
            this.id = row.ItemArray[4].ToString();
          
            foreach (string str in co)
            {
                authors.Add(str);
            }
            this.name = row.ItemArray[7].ToString();
            this.year = row.ItemArray[8].ToString();
            this.link = row.ItemArray[11].ToString();
            this.extra_inf = row.ItemArray[12].ToString();
            this.city = row.ItemArray[23].ToString();
            this.publish_house = row.ItemArray[24].ToString();
            this.category = row.ItemArray[6].ToString();
            this.label = row.ItemArray[10].ToString();
            this.full_pages_number = row.ItemArray[9].ToString();
            this.first_page = row.ItemArray[16].ToString();
            this.last_page = row.ItemArray[17].ToString();
            this.part = row.ItemArray[18].ToString();
            this.name_of_collection = row.ItemArray[22].ToString();
            this.short_name_conf = row.ItemArray[21].ToString();
            this.place_conf = row.ItemArray[25].ToString();
            this.date_conf = row.ItemArray[26].ToString();
        }
        public override Excel.Worksheet Excel(Excel.Worksheet excelws, int i)
        {
            string auth = "";
            foreach (string tez in this.authors)
            {
                //тут
                auth += ", ";
                auth += tez;
            }
            auth = auth.Substring(2, auth.Length - 2);

            Excel.Range excelcells = excelws.Cells[i, 1];
            excelcells.Value = auth;

            excelcells = excelws.Cells[i, 2];
            excelcells.Value = this.name;

            excelcells = excelws.Cells[i, 3];
            excelcells.Value = this.year;

            excelcells = excelws.Cells[i, 4];
            excelcells.Value = this.full_pages_number;

            excelcells = excelws.Cells[i, 5];
            excelcells.Value = this.label;

            excelcells = excelws.Cells[i, 6];
            excelcells.Value = this.link;

            excelcells = excelws.Cells[i, 7];
            excelcells.Value = this.extra_inf;

            excelcells = excelws.Cells[i, 8];
            excelcells.Value = this.city;

            excelcells = excelws.Cells[i, 9];
            excelcells.Value = this.publish_house;

            excelcells = excelws.Cells[i, 10];
            excelcells.Value = this.first_page;
            excelcells = excelws.Cells[i, 11];
            excelcells.Value = this.last_page;
            excelcells = excelws.Cells[i, 12];
            excelcells.Value = this.part;
            excelcells = excelws.Cells[i, 13];
            excelcells.Value = this.name_of_collection;
            excelcells = excelws.Cells[i, 14];
            excelcells.Value = this.short_name_conf;
            excelcells = excelws.Cells[i, 15];
            excelcells.Value = this.place_conf;
            excelcells = excelws.Cells[i, 16];
            excelcells.Value = this.date_conf;
            return excelws;
        }
    }

    public class Otchet : Book
    {
        public string label;
        public string inv_number;
        public string registr_number;
        public string date_registr;
        public string status;
        public string theme;
        public string full_pages_number;
        // public string publish_house;
        public override string list_Word()
        {
            string result = "";
            foreach (string author in this.authors)
            {
                result = result + author + ", ";// later make so, that it wouldnt print "," after last author;
            }
            int len = result.Length;
            result = result.Substring(0, len - 2);
            result = result + " Отчет о научно-исследователькой работе «" + this.name + "»";
            result = result + "(" + this.status + ")";
            result = result + ", инвентарный №" + this.registr_number + " от " + this.date_registr + ", ";
            result = result + " по теме «" + this.theme + "», регистрационный № " + this.registr_number + ". ";
            result = result + this.city + ": " + this.publish_house + ", " + this.year + ". ";
            result = result + this.full_pages_number + " c. (Депонировано в ЦИТиС). ";
            if (this.link != null)
            {
                result = result + "URL: " + this.link + ". ";
            }
            if (this.extra_inf != null)
            {
                result = result + "(" + this.extra_inf + ") ";
            }
            return result;
        }
        public override string[] table_Word(string name_a)
        {
            string[] elems = new string[5];
            elems[0] = this.name + "(" + this.category + ")";
            elems[1] = this.label;
            string result = "";
            result = result + "Отчет о научно-исследователькой работе «" + this.name + "»";
            result = result + "(" + this.status + ")";
            result = result + ", инвентарный №" + this.registr_number + " от " + this.date_registr + ", ";
            result = result + " по теме «" + this.theme + "», регистрационный № " + this.registr_number + ". ";
            result = result + this.city + ": " + this.publish_house + ", " + this.year + ". ";
            result = result + this.full_pages_number + " c. (Депонировано в ЦИТиС). ";
            if (this.link != null)
            {
                result = result + "URL: " + this.link + ". ";
            }
            if (this.extra_inf != null)
            {
                result = result + "(" + this.extra_inf + ") ";
            }

            elems[2] = result;
            elems[3] = this.full_pages_number;
            elems[4] = "";
            foreach (string auth in this.authors)
            {
                if (!auth.Equals(name_a))
                {
                    elems[4] = auth + ", ";
                }
            }
     
            return elems;
          
        }
        public override void fill(DataRow row, List<string> co)
        {
            

            //this.authors.Add(row.ItemArray[1].ToString());
            this.id = row.ItemArray[4].ToString();
            this.name = row.ItemArray[7].ToString();
            
            foreach (string str in co)
            {

                this.authors.Add(str);
            }
            this.year = row.ItemArray[8].ToString();
            this.link = row.ItemArray[11].ToString();
            this.extra_inf = row.ItemArray[12].ToString();
            this.city = row.ItemArray[28].ToString();
            this.publish_house = row.ItemArray[29].ToString();
            this.category = row.ItemArray[6].ToString();
            this.label = row.ItemArray[10].ToString();
            this.inv_number = row.ItemArray[30].ToString();
            this.registr_number = row.ItemArray[31].ToString();
            this.date_registr = row.ItemArray[32].ToString();
            this.status = row.ItemArray[33].ToString();
            this.theme = row.ItemArray[34].ToString();
            this.full_pages_number = row.ItemArray[9].ToString();
            //return report;
        }
        public override Excel.Worksheet Excel(Excel.Worksheet excelws, int i)
        {
            string auth = "";
            //int k = 0;
            //int j = 0;
            //foreach (string tez in this.authors)
            //{
            //    foreach (string tez1 in this.authors)
            //    {
            //        if (String.Compare(tez, tez1) == 0 && j != k)
            //        {
            //            string ex = tez;
            //            authors.;
            //        }
            //        j++;
            //    }
            //    k++;
            //}
            foreach (string tez in this.authors)
            {
                auth += ", ";
                auth += tez;
            }
            auth = auth.Substring(2, auth.Length - 2);

            Excel.Range excelcells = excelws.Cells[i, 1];
            excelcells.Value = auth;

            excelcells = excelws.Cells[i, 2];
            excelcells.Value = this.name;

            excelcells = excelws.Cells[i, 3];
            excelcells.Value = this.year;

            excelcells = excelws.Cells[i, 4];
            excelcells.Value = this.full_pages_number;

            excelcells = excelws.Cells[i, 5];
            excelcells.Value = this.label;

            excelcells = excelws.Cells[i, 6];
            excelcells.Value = this.link;

            excelcells = excelws.Cells[i, 7];
            excelcells.Value = this.extra_inf;

            excelcells = excelws.Cells[i, 8];
            excelcells.Value = this.city;

            excelcells = excelws.Cells[i, 9];
            excelcells.Value = this.publish_house;

            excelcells = excelws.Cells[i, 10];
            excelcells.Value = this.inv_number;

            excelcells = excelws.Cells[i, 11];
            excelcells.Value = this.registr_number;

            excelcells = excelws.Cells[i, 12];
            excelcells.Value = this.date_registr;

            excelcells = excelws.Cells[i, 13];
            excelcells.Value = this.theme;

            excelcells = excelws.Cells[i, 14];
            excelcells.Value = this.status;

            return excelws;
        }
    }
    public class Preprint : Book
    {
        public string label;
        public string full_pages_number;
        public override string list_Word()
        {
            string result = "";
            foreach (string author in this.authors)
            {
                result = result + author + ", ";// later make so, that it wouldnt print "," after last author;
            }
            int len = result.Length;
            result = result.Substring(0, len - 2);
            
            result = result + " " + this.name + ". ";
            result = result + this.year + ". ";
            result = result + this.full_pages_number + "c. ";
            if (this.link != null)
            {
                result = result + "URL: " + this.link + ". ";
            }
            if (this.extra_inf != null)
            {
                result = result + "(" + this.extra_inf + ") ";
            }
            return result;
        }
        public override string[] table_Word(string name_a)
        {
            string[] elems = new string[5];
            elems[0] = this.name + "(" + this.category + ")";
            elems[1] = this.label;
            string result = "";
            result = result + this.name + ". ";
            result = result + this.year + ". ";
            result = result + this.full_pages_number + "c. ";
            if (this.link != null)
            {
                result = result + "URL: " + this.link + ". ";
            }
            if (this.extra_inf != null)
            {
                result = result + "(" + this.extra_inf + ") ";
            }
            elems[2] = result;
            elems[3] = this.full_pages_number;
            elems[4] = "";
            foreach (string auth in this.authors)
            {
                if (!auth.Equals(name_a))
                {
                    elems[4] = auth + ", ";
                }
            }

            return elems;
        }
        public override void fill(DataRow row, List<string> co)
        {

            this.id = row.ItemArray[4].ToString();
            
            
            foreach (string str in co)
            {

                this.authors.Add(str);
            }
            this.name = row.ItemArray[7].ToString();
            this.year = row.ItemArray[8].ToString();
            this.link = row.ItemArray[11].ToString();
            this.extra_inf = row.ItemArray[12].ToString();
            this.city = row.ItemArray[28].ToString();
            this.category = "препринт";
            this.label = row.ItemArray[10].ToString();
            this.full_pages_number = row.ItemArray[9].ToString();
            //return preprint;
        }
        public override Excel.Worksheet Excel(Excel.Worksheet excelws, int i)
        {
            string auth = "";
            foreach (string tez in this.authors)
            {
                auth += ", ";
                auth += tez;
            }
            auth = auth.Substring(2, auth.Length - 2);

            Excel.Range excelcells = excelws.Cells[i, 1];
            excelcells.Value = auth;

            excelcells = excelws.Cells[i, 2];
            excelcells.Value = this.name;

            excelcells = excelws.Cells[i, 3];
            excelcells.Value = this.year;

            excelcells = excelws.Cells[i, 4];
            excelcells.Value = this.full_pages_number;

            excelcells = excelws.Cells[i, 5];
            excelcells.Value = this.label;

            excelcells = excelws.Cells[i, 6];
            excelcells.Value = this.link;

            excelcells = excelws.Cells[i, 7];
            excelcells.Value = this.extra_inf;
            return excelws;
        }
    }


    public class TexThesis
    {

        private const int ALPHABET_LEN = 30;

        public StreamWriter output;
        private string filename;
        private string[] lat = {  "A",  "B",    "V", "G",   "D",  "E",
                             "YO", "ZH",    "Z", "I",  "IY",  "K",
                              "L",  "M",    "N", "O",   "P",  "R",
                              "S",  "T",    "U", "F",  "KH", "TS",
                             "CH", "SH", "SHCH", "EH", "YU", "YA",
                           };

        private string[] kir = { "А", "Б", "В", "Г", "Д", "Е",
                             "Ё", "Ж", "З", "И", "Й", "К",
                             "Л", "М", "Н", "О", "П", "Р",
                             "С", "Т", "У", "Ф", "Х", "Ц",
                             "Ч", "Ш", "Щ", "Э", "Ю", "Я",
                           };

        public TexThesis(string filename)
        {
            this.filename = filename;
        }

        private bool OutputIsOpen
        {
            get
            {
                return output != null && output.BaseStream != null;
            }
        }

        public void writeBibliographyThesis(List<Book> lBook)
        {
            Random rnd = new Random();
            int r = rnd.Next(1, 100);
            string localFilename = filename + "\\" + "tex" + r + ".txt";
            using (output = OutputIsOpen ? output : new StreamWriter(localFilename))
            {
                output.WriteLine("\\begin{thebibliography}{00.}");
                foreach (Book book in lBook)
                {
                    Write(book, true);
                }
                output.Write("\\end{thebibliography}");
                MessageBox.Show("TeX file Added ");
            }
        }


        public void Write(Book book, bool close)
        {
            try
            {
                output = OutputIsOpen ? output : new StreamWriter(this.filename);
                output.Write("\\bibitem" + "{" + getKey(book.authors) + getYear(book) + "}" + getAuthors(book.authors) + " " + getName(book));

                if (book is Tezis)
                {
                    Tezis tezis = (Tezis)book;
                    output.Write(" // " + getNameOfCollection(tezis) + getPart(tezis) + " " + getPlace(tezis) + " ");
                }

                if (book is Otchet)
                {
                    Otchet otchet = (Otchet)book;
                    output.Write("." + getNameAndInfo(otchet));
                }

                output.Write(getCityPubHouseYear(book) + " ");

                if (book is Tezis)
                {
                    Tezis tezis = (Tezis)book;
                    output.Write(getPage(tezis) + " ");
                }

                if (book is Otchet)
                {
                    Otchet otchet = (Otchet)book;
                    output.Write(getFullPage(otchet));
                }

                if (book is Preprint)
                {
                    Preprint preprint = (Preprint)book;
                    output.Write(getFullPage(preprint));
                }

                output.WriteLine(getLinkAndInfo(book) + "\n");
            }
            catch (IOException e)
            {
            }
            finally
            {
                if (!close)
                {
                    output.Close();
                }
            }

        }

        private string getKey(List<string> authors)
        {

            string result = "";

            foreach (string str in authors)
            {
                string res = str.Substring(0, Math.Min(3, str.Length)).ToUpper();
                string test = res;

                for (int i = 0; i < 30; i++)
                {
                    res = res.Replace(kir[i], lat[i]);
                }

                result += res;
            }

            return result;
        }

        private string getYear(Book book)
        {
            return book.year;
        }

        private String getAuthors(List<string> authors)
        {
            string result = "";
            foreach (string str in authors)
            {
                string res = str;
                res = res.Replace(" ", "~");
                //res = res.Replace(".", ".~");
                res = res.Substring(0, res.Length - 1) + "., ";
                result += res;
            }
            result = " \\emph{" + result.Substring(0, result.Length - 2) + " \\/}";
            return result;
        }

        private string getName(Book book)
        {
            return book.name;
        }

        private string getNameOfCollection(Tezis tezis)
        {
            return tezis.name_of_collection;
        }

        private string getPart(Tezis tezis)
        {
            string result = "";
            result += "T.~" + tezis.part + ".";
            return result;
        }

        private string getPlace(Tezis tezis)
        {
            string result = "";

            if (tezis.place_conf != null)
            {
                result += tezis.place_conf + ", ";
            }

            result += tezis.city + ", ";

            if (tezis.date_conf != null)
            {
                string data = tezis.date_conf;
                data = data.Replace("-", "--");
                data = data.Replace("г.", "~г.");
                result += data;
            }
            else
            {
                result = result.Substring(0, result.Length - 2);
            }

            result = "(" + result + ")";

            return result;
        }

        private string getCityPubHouseYear(Book book)
        {
            string result = "";
            if (book.city != null && book.publish_house != null && book.year != null)
                result = this.getShortCityname(book.city) + ".: " + book.publish_house + ", " + book.year + ".";
            return result;
        }

        private string getPage(Tezis tezis)
        {
            string result = "";
            result = "c.~";
            int firstPage = 0, lastPage = 0;

            try
            {
                firstPage = Convert.ToInt32(tezis.first_page);
                lastPage = Convert.ToInt32(tezis.last_page);
            }
            catch (FormatException e)
            {
                Console.Write("NumberFormat Exception");
            }

            if (firstPage > lastPage)
            {
                result += tezis.first_page + "--" + tezis.last_page;
            }
            else
            {
                result += tezis.first_page;
            }
            return result;
        }

        private string getLinkAndInfo(Book book)
        {
            string result = "";

            if (book.link != null)
            {
                result += book.link + ". ";
            }

            if (book.extra_inf != null)
            {
                string inf = "( " + book.extra_inf + " )";
                result += inf;
            }

            return result;
        }

        private string getShortCityname(string str)
        {
            if (str.Equals("Москва"))
            {
                return "М";
            }
            else
            {
                return "СПб";
            }
        }

        private string getNameAndInfo(Otchet otchet)
        {
            return " Отчет о научно-исследователькой работе " + "<<" + otchet.name + ">> " + otchet.status + " инвертарный " + "\\textnumero~" + otchet.inv_number +
               " от " + otchet.date_registr + ", " + "<<" + otchet.theme + ">> " + "регистрационный " + " \\textnumero~" + otchet.registr_number + ". ";
        }

        private string getFullPage(Otchet otchet)
        {
            return otchet.full_pages_number + "~c." + " (Депонировано в ЦИТиС.)";
        }

        private string getFullPage(Preprint preprint)
        {
            return preprint.full_pages_number + "~c.";
        }
    }
}
