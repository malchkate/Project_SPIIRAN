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
    public partial class ReportForm : Form
    {
        OleDbConnection connection;
        OleDbDataAdapter dataAdapter;
        int numToAdd;
        public ReportForm(OleDbConnection connection, OleDbDataAdapter dataAdapter, string dbPath, int numToAdd)
        {
            this.connection = connection;
            this.dataAdapter = dataAdapter;
            this.numToAdd = numToAdd;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string str1 = textBox1.Text;
            string str2 = textBox2.Text;
            string str3 = textBox3.Text;
            string str4 = textBox4.Text;
            string str5 = textBox5.Text;
            string str6 = comboBox1.Text;
            string str7 = textBox6.Text;
            try
            {
                // открытие соединения
                connection.Open();
                // создание команды, соответствующей соединению
                OleDbCommand commandUpd = connection.CreateCommand();
                // текст запроса на вставку
                string textQuery = @"INSERT INTO Reports (id, city, publishing, inventory_num, reg_num, data, status, theme)" +
                                   " VALUES (" + numToAdd + ", '" + str1 + "', '" + str2 + "', '" + str3 + "', '" + str4 + "', '" + str5 + "', '" + str6 +"', '" + str7 + "')";
                commandUpd.CommandText = textQuery;
                // выполнение запроса
                commandUpd.ExecuteNonQuery();
                // закрытие соединения
                connection.Close();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error!:" + ex.Message + ex.StackTrace);
            }
        }
    }
}
