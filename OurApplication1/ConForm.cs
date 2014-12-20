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
    public partial class ConForm : Form
    {
        OleDbConnection connection;
        OleDbDataAdapter dataAdapter;
        DataTable table;
        DataSet ds;
        Label authorLabel;
        int numToAdd;
        public ConForm(OleDbConnection connection, OleDbDataAdapter dataAdapter, string dbPath, Label label, int numToAdd)
        {
            this.connection = connection;
            this.dataAdapter = dataAdapter;
            this.authorLabel = label;
            this.numToAdd = numToAdd;
            InitializeComponent();
            try
            {
                string queryString = "SELECT * FROM Conferences";
                string conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbPath;
                connection.ConnectionString = conStr;
                dataAdapter = new OleDbDataAdapter(queryString, connection);
                ds = new DataSet();
                dataAdapter.Fill(ds, "Conferences");
                table = ds.Tables["Conferences"];
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

        private void button2_Click(object sender, EventArgs e)
        {
            authorLabel.Text = listBox.SelectedItem.ToString();

            string queryString = "SELECT * FROM Attributes";
            dataAdapter = new OleDbDataAdapter(queryString, connection);
            ds = new DataSet();
            dataAdapter.Fill(ds, "Attributes");
            table = ds.Tables["Attributes"];
            string str6 = textBox6.Text;
            string str7 = textBox7.Text;
            string str8 = textBox8.Text;
            int id = 0;
            foreach (DataRow row in table.Rows)
            {
                if (id < Convert.ToInt32(row[0]))
                    id = Convert.ToInt32(row[0]);
            }
            id += 1;

            queryString = "SELECT * FROM Conferences";
            dataAdapter = new OleDbDataAdapter(queryString, connection);
            ds = new DataSet();
            dataAdapter.Fill(ds, "Conferences");
            table = ds.Tables["Conferences"];
            int conId = 5;
            foreach (DataRow row in table.Rows)
            {
                if (String.Compare(row[1].ToString(), listBox.SelectedItem.ToString()) == 0)
                {
                    conId = Convert.ToInt32(row[0]);
                    break;
                }
            }
            try
            {
                connection.Open();
                OleDbCommand commandUpd = connection.CreateCommand();
                string textQuery = "INSERT INTO Attributes (id, first_page, last_page, volume, conference_id)" +
                                   " VALUES (" + numToAdd + ", '" + str6 + "', '" + str7 + "', '" + str8 + "', " + conId + ")";
                commandUpd.CommandText = textQuery;
                commandUpd.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error!:" + ex.Message + ex.StackTrace);
            }
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string str0 = textBox0.Text;
            string str1 = textBox1.Text;
            string str2 = textBox2.Text;
            string str3 = textBox3.Text;
            string str4 = textBox4.Text;
            string str5 = textBox5.Text;
            
            try
            {
                connection.Open();
                OleDbCommand commandUpd = connection.CreateCommand();
                string textQuery = "INSERT INTO Conferences (id, conf_name, book_name, city, editors, place, dates)" +
                                   " VALUES (" + numToAdd + ", '" + str0 + "', '" + str1 + "', '" + str2 + "', '" + str3 + "', '" + str4 + "', '" + str5 + "' )";
                commandUpd.CommandText = textQuery;
                commandUpd.ExecuteNonQuery();
                listBox.Items.Add(str0);
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error!:" + ex.Message);
            }
        }
    }
}
