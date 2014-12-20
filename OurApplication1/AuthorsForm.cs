using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OurApplication1
{
    public partial class AuthorsForm : Form
    {
        public string[] authorsLst;
        OleDbConnection connection;
        OleDbDataAdapter dataAdapter;
        DataTable table;
        DataSet ds;
        Label authorLabel;
        int[] num;
        public AuthorsForm(OleDbConnection connection, string dbPath, Label label, ref int[] num)
        {
            authorsLst = new string[100];
            this.connection = connection;
            this.authorLabel = label;
            this.num = num;
            InitializeComponent();
            try
            {
                string queryString = "SELECT * FROM Authors";
                string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath;
                connection.ConnectionString = conStr;
                dataAdapter = new OleDbDataAdapter(queryString, connection);
                ds = new DataSet();
                dataAdapter.Fill(ds, "Authors");
                table = ds.Tables["Authors"];
                foreach (DataRow row in table.Rows)
                {
                    listBox.Items.Add(row[1]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error!:" + ex.Message + ex.StackTrace);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string author = authorsBox.Text;
            int code = 0;
            foreach (DataRow row in table.Rows)
            {
                if (System.String.Compare(row[1].ToString(), author) == 0)
                {
                    errorBox.Text = "Такой автор уже есть в базе данных!";
                    return;
                }
                if (code < System.Convert.ToInt32(row[0]))
                    code = System.Convert.ToInt32(row[0]);
            }
            int num = code + 1;
            try
            {
                // открытие соединения
                connection.Open();
                // создание команды, соответствующей соединению
                OleDbCommand commandUpd = connection.CreateCommand();
                // текст запроса на вставку
                string textQuery = "INSERT INTO Authors (id, author)" +
                                   " VALUES (" + num + ", '" + author + "' )";
                commandUpd.CommandText = textQuery;
                // выполнение запроса
                commandUpd.ExecuteNonQuery();
                listBox.Items.Add(author);
                // закрытие соединения
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error!:" + ex.Message + ex.StackTrace);
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            int i = 0;
            string str = listBox.SelectedItems[0].ToString();
            string authors = "";

            dataAdapter.Fill(ds, "Authors");
            table = ds.Tables["Authors"];
            
            while (str != null)
            {
                foreach (DataRow row in table.Rows)
                {
                    if (row[1].ToString() == str)
                    {
                        num[i] = new int();
                        num[i] = Convert.ToInt32(row[0]);
                        break;
                    }
                }
                authors += str + ", ";
                i++;
                try
                {
                    str = listBox.SelectedItems[i].ToString();
                }
                catch (IndexOutOfRangeException)
                {
                    str = null;
                }
            }
            authorLabel.Text = authors;
            this.Close();
        }
    }
}
