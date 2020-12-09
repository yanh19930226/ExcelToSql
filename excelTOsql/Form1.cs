using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Data.OleDb;
namespace excelTOsql
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OleDbConnection oledbcConnection;
        string connString="server=localhost;uid=sa;pwd=123;database=TTT";//这里是自己的数据库信息，根据自己的情况修改
        /*打开excel文件并选择表单*/
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openDG = new OpenFileDialog();
                openDG.Title = "打开Excel表格";
                //openDG.Filter = "Excel表格(*.xlsx)|*.xlsx|CSV格式(*.csv)|*.csv|所有文件(*.*)|*.*";
                openDG.ShowDialog();
                string filename;//文件路径
                filename = openDG.FileName;
                filenamebox.Text = filename;//显示文件路径
                string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";Extended Properties='Excel 12.0;HDR=False;IMEX=1'";//此处为对excel的读取设置，不同的excel版本有不同的设定。
                
                
                //@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + ";Extended Properties=Excel 8.0";
               /*读取excel数据到内存中*/
                oledbcConnection = new OleDbConnection(strConn);
                oledbcConnection.Open();
                DataTable table = new DataTable();
                table = oledbcConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                /*获取此文件中的所有表单*/
                excelBookComboBoxEx.Items.Clear();
                foreach (DataRow dr in table.Rows)
                {
                    excelBookComboBoxEx.Items.Add((String)dr["TABLE_NAME"]);
                }
                excelBookComboBoxEx.Text = excelBookComboBoxEx.Items[0].ToString();

                DataSet ds = new DataSet();
                SqlConnection conn = new SqlConnection(connString);

                string strConn1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filenamebox.Text + ";Extended Properties='Excel 12.0;HDR=False;IMEX=1'";// "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=False;IMEX=1'";
                //"Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + excelFile + ";" + "Extended Properties=Excel 8.0;";
                OleDbConnection comm = new OleDbConnection(strConn1);
                comm.Open();

                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                strExcel = string.Format("select * from [{0}]", excelBookComboBoxEx.Text);
                myCommand = new OleDbDataAdapter(strExcel, strConn1);
                myCommand.Fill(ds, excelBookComboBoxEx.Text);
                //filenamebox.Text = ds.Tables[0].Rows.Count.ToString();/*显示导入数据的总条数*/
                comm.Close();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        public void TransferData(string excelFile, string sheetName, string connectionString)
        {
            DataSet ds = new DataSet();
            try
            {
                //获取全部数据
                string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFile + ";Extended Properties='Excel 12.0;HDR=False;IMEX=1'";// "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=False;IMEX=1'";
                //"Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + excelFile + ";" + "Extended Properties=Excel 8.0;";
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                strExcel = string.Format("select * from [{0}]", sheetName);
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                myCommand.Fill(ds, sheetName);
                //如果目标表不存在则创建
                string strSql = string.Format("if object_id('{0}') is null create table {0}(", sheetName);
                foreach (System.Data.DataColumn c in ds.Tables[0].Columns)
                {
                    strSql += string.Format("[{0}] varchar(255),", c.ColumnName);
                }
                strSql = strSql.Trim(',') + ")";
                using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(connectionString))
                {
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;
                    command.ExecuteNonQuery();
                    sqlconn.Close();
                }
                //用bcp导入数据
                using (System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(connectionString))
                {
                    bcp.BatchSize = 100;//每次传输的行数
                    bcp.NotifyAfter = 100;//进度提示的行数
                    bcp.DestinationTableName = sheetName;//目标表
                    bcp.WriteToServer(ds.Tables[0]);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string FILE_NAME = filenamebox.Text;      
            string ST = excelBookComboBoxEx.Text;
            TransferData(FILE_NAME, ST, connString);
            MessageBox.Show("导入成功!");
        }     
    }
}
