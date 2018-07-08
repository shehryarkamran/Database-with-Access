using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb; // to connect old db

namespace Database_with_Access
{
    public partial class Form1 : Form
    {
        OleDbConnection connect = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source= office.accdb;Persist Security Info=False;");
        OleDbCommand command = new OleDbCommand();
        public Form1()
        {
            InitializeComponent();
        }
        // insert btn
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                connect.Open();
                command.Connection = connect;
                command.CommandText = "INSERT INTO employee (Id, Emp_Name,salary) values ('" + textBox1.Text + "','" + this.textBox2.Text + "',' "+textBox3.Text +" ')";
                OleDbDataReader reader = command.ExecuteReader();
                MessageBox.Show("Data Saved");
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
            }
            catch (Exception a)
            {
                MessageBox.Show("error", a.Message);
            }
            finally
            {
                connect.Close();
                ftn();
            }
        }
        //update
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                connect.Open();
                command.Connection = connect;
                command.CommandText = "UPDATE Employee set Emp_name = '" + textBox2.Text + "', Salary = '" + textBox3.Text+ "' where ID =" +int.Parse( textBox1.Text);
                OleDbDataReader reader = command.ExecuteReader();
                MessageBox.Show("Data updated");
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
            } catch(Exception a)
            {
                MessageBox.Show(a.Message, "error");
            }
            finally
            {
                connect.Close();
                ftn();
            }
        }
            private void textBox1_TextChanged(object sender, EventArgs e)
        {}
        //delete btn
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                connect.Open();
                command.Connection = connect;
                    command.CommandText = "DELETE FROM employee where ID ="  +int.Parse( textBox1.Text);
                    OleDbDataReader reader =
                   command.ExecuteReader();
                    MessageBox.Show("Data Deleted");
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
            } catch(Exception a)
            {
                MessageBox.Show(a.Message,"error");
            }
            finally
            {
                connect.Close();
                ftn();
            }
            //reader.Close();
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {}
        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.Columns.Add("ID", "Employee ID");
            dataGridView1.Columns.Add("EMP_NAME", "Name");
            dataGridView1.Columns.Add("SALARY", "Salary");
            ftn();
        }
        //retrive btn
        private void ftn()
        {
            dataGridView1.Rows.Clear();
            try
            {
                connect.Open();
                command.Connection = connect;
                command.CommandText = "select * from Employee";
                OleDbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells["ID"].Value = reader[0].ToString();
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells["EMP_NAME"].Value = reader[1].ToString();
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells["SALARY"].Value = reader[2].ToString();
                    //dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells["empdept"].Value = reader[3].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("error", ex.Message);
            }
            finally
            {
                connect.Close();
            }
        }
    }
}