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

namespace EXCEL_COMPARISON
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "*.xlsx|*.xlsx";
            openFileDialog1.ShowDialog();
            string filename = openFileDialog1.FileName;
            Read_Excel(filename, dataGridView1, textBox1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "*.xlsx|*.xlsx";
            openFileDialog1.ShowDialog();
            string filename = openFileDialog1.FileName;
            Read_Excel(filename, dataGridView2, textBox2);
            
        }

        private void Read_Excel(string filename, DataGridView dgv, TextBox TextBox)
        {
            DataTable ExcelTable;
            DataSet ds = new DataSet();
            TextBox.Text = filename;

            //Excel的连接
			//connect to excel
            string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + filename + ";Extended Properties='Excel 12.0; HDR=NO; IMEX=1'";
            OleDbConnection objConn = new OleDbConnection(strConn);

           try
            {
                objConn.Open();
                DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                string tableName = schemaTable.Rows[0][2].ToString().Trim();//获取 Excel 的表名，默认值是sheet1 // get the name of excel
                string strSql = "select * from [" + tableName + "]";
                OleDbCommand objCmd = new OleDbCommand(strSql, objConn);
                OleDbDataAdapter myData = new OleDbDataAdapter(strSql, objConn);
                myData.Fill(ds, tableName);//填充数据 //fill data

                ExcelTable = ds.Tables[tableName];
                int iColums = ExcelTable.Columns.Count;//列数 // colnum
                int iRows = ExcelTable.Rows.Count;//行数 //rownum


                //定义二维数组存储 Excel 表中读取的数据
				//define the array which is used to store data in excel
                string[,] storedata = new string[iRows, iColums];

                for (int i = 0; i < ExcelTable.Rows.Count; i++)
                {
                    for (int j = 0; j < ExcelTable.Columns.Count; j++)
                    {
                        //将Excel表中的数据存储到数组
						//push data in the excel to array
                        storedata[i, j] = ExcelTable.Rows[i][j].ToString();
                    }
                }

                dgv.DataSource = ExcelTable;
                objConn.Close();
           }

            catch (Exception)
            {
                MessageBox.Show("Error!");
            }
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            DataGridViewTextBoxColumn dgv_Text = new DataGridViewTextBoxColumn();
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                int j = i + 1;
                dataGridView1.Rows[i].HeaderCell.Value = j.ToString();
            }
        }

        private void dataGridView2_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            DataGridViewTextBoxColumn dgv_Text = new DataGridViewTextBoxColumn();
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                int j = i + 1;
                dataGridView2.Rows[i].HeaderCell.Value = j.ToString();
            }
        }
    }
}
