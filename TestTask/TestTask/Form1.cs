using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using TestTask.Models;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestTask
{
    public partial class Form1 : Form
    {
        private SqlConnection _sqlConnection = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox2.Text = comboBox2.Items[1].ToString();
            _sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["TestDB"].ConnectionString);

            _sqlConnection.Open();

            SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM Employee", _sqlConnection);

            DataSet db = new DataSet();
            adapter.Fill(db);
            dataGridView1.DataSource = db.Tables[0];

        }
        private void button1_Click(object sender, EventArgs e)
        {
            Employee emp = new Employee(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, DateTime.Parse(textBox6.Text), DateTime.MinValue, StateRecord.Active);
            emp.DbInsert(textBox7.Text, _sqlConnection);
            MessageBox.Show("Принят");
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT * FROM Employee", _sqlConnection);

            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0];
        }


        private void RefreshGrid()
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT * FROM Employee", _sqlConnection);

            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0];
        }
        private void button1_Click_2(object sender, EventArgs e)
        {
            RefreshGrid();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            string str = "";
            if (textBox9.Text.Length > 0)
            {
                str = $"Number LIKE '%{textBox9.Text}%'";
            }
            if (textBox8.Text.Length > 0)
            {
                if (textBox9.Text.Length > 0)
                {
                    str += " AND ";
                }
                str += $"Name LIKE '%{textBox8.Text}%'";
            }
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = str;
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlDataAdapter adapter = new SqlDataAdapter($"SELECT Employee.Name, Employee.Number, Employee.JobTitle, SubDivision.NameDepartment, Employee.Email, Employee.PhoneNumber, Employee.HireDate, Employee.DismissDate FROM Employee, SubDivision WHERE Employee.ID_SubDivision = SubDivision.Id", _sqlConnection);

            DataSet db = new DataSet();
            adapter.Fill(db);
            dataGridView1.DataSource = db.Tables[0];
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = $"NameDepartment LIKE '{comboBox1.Text}'";
        }

        private void comboBox1_MouseDown(object sender, MouseEventArgs e)
        {
            comboBox1.Items.Clear();
            SqlCommand command = new SqlCommand("SELECT * FROM SubDivision", _sqlConnection);
            SqlDataReader sqlDataReader = command.ExecuteReader();
            List<object> list = new List<object>();
            while (sqlDataReader.Read())
            {
                list.Add(sqlDataReader["NameDepartment"]);
            }
            comboBox1.Items.AddRange(list.ToArray());
            sqlDataReader.Close();
        }

        private void DataFilter()
        {
            DateTime dateTime1 = dateTimePicker1.Value;
            DateTime dateTime2 = dateTimePicker2.Value;
            string str;
            if (comboBox2.Text == comboBox2.Items[0].ToString())
            {
                str = "DismissDate";
            }
            else
            {
                str = "HireDate";
            }
            SqlDataAdapter dataAdapter = new SqlDataAdapter($"SELECT * FROM Employee WHERE {str} >= '{dateTime1.Month}/{dateTime1.Day}/{dateTime1.Year}' AND {str} <= '{dateTime2.Month}/{dateTime2.Day}/{dateTime2.Year}'", _sqlConnection);

            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0];
        }
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DataFilter();
        }


        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            DataFilter();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if((dataGridView1.DataSource as DataTable).DefaultView.Count == 1)
            {
                DataTableReader reader = (dataGridView1.DataSource as DataTable).DefaultView.ToTable().CreateDataReader();
                int id = 0;
                if (reader.Read())
                {
                    id = (int)reader["Id"];
                }
                SqlCommand cmd = new SqlCommand("UPDATE Employee SET DismissDate = @DismissDate WHERE Id = @Id", _sqlConnection);
                cmd.Parameters.AddWithValue("@DismissDate", $"{DateTime.Now.Month}/{DateTime.Now.Day}/{DateTime.Now.Year}");
                cmd.Parameters.AddWithValue("@Id", id);

                cmd.ExecuteNonQuery();
                RefreshGrid();
            }
            else
            {
                MessageBox.Show("Выберите одного сотрудника!");
            }
        }
        private void panel1_DragDrop(object sender, DragEventArgs e)
        {
            label13.Text = "Перетащите файл сюда";

            string[] link = (string[])e.Data.GetData(DataFormats.FileDrop);
            label14.Text = link[0];

        }

        private void panel1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
                label13.Text = "Отпустите файл!";
            }
        }

        private void panel1_DragLeave(object sender, EventArgs e)
        {
            label13.Text = "Перетащите файл сюда";
        }

        private void button3_Click(object sender, EventArgs e)
        {

            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb;
            Excel.Worksheet ws;
            wb = excel.Workbooks.Open(label14.Text);
            ws = wb.Worksheets[2];
            int i = 2;
            int j = 1;
            List<string> list = new List<string>();
            SubDivision subD;
            while (Convert.ToString(ws.Cells[i, j].Value) != null)
            {
                while (j < 4)
                {
                    list.Add(Convert.ToString(ws.Cells[i, j].Value));
                    j++;
                }
                subD = new SubDivision(list[1], list[2] == null ? "" : list[2]);
                list = new List<string>();
                subD.DbInsert(_sqlConnection);
                j = 1;
                i++;
            }
            i = 2;
            j = 1;
            ws = wb.Worksheets[1];
            list = new List<string>();
            Employee emp;
            while (Convert.ToString(ws.Cells[i,j].Value) != null)
            {
                while (j < 10)
                {
                    list.Add(Convert.ToString(ws.Cells[i,j].Value));
                    j++;
                }
                emp = new Employee(list[1], list[2], list[3], list[5], list[6], DateTime.Parse(list[7]), list[8] == null ? DateTime.MinValue : DateTime.Parse(list[8]), StateRecord.Active);
                emp.DbInsert(list[4], _sqlConnection);
                list = new List<string>();
                j = 1;
                i++;
            }
            wb.Close();
            MessageBox.Show("Выполнено");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if ((dataGridView1.DataSource as DataTable).DefaultView.Count == 1)
            {
                DataTableReader reader = (dataGridView1.DataSource as DataTable).DefaultView.ToTable().CreateDataReader();
                int id = 0;
                if (reader.Read())
                {
                    id = (int)reader["Id"];
                }
                SqlCommand cmd = new SqlCommand("DELETE FROM Employee WHERE Id = @Id", _sqlConnection);
                cmd.Parameters.AddWithValue("@Id", id);

                cmd.ExecuteNonQuery();
                RefreshGrid();
            }
            else
            {
                MessageBox.Show("Выберите одну запись!");
            }
        }
    }
}
