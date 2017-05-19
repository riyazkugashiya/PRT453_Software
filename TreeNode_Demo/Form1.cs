using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using STATCONNECTORCLNTLib;
using StatConnectorCommonLib;
using STATCONNECTORSRVLib;
using System.IO;
using System.Data.OleDb;

namespace TreeNode_Demo
{
    public partial class Form1 : Form
    {
        private string Excel03conString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0{1}'";
        private string Excel07conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0{1}'";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //object o1;
            //int n = 20;
            //StatConnector sc1 = new STATCONNECTORSRVLib.StatConnector();
            //sc1.Init("R");
            //sc1.SetSymbol("n1", n);
            //sc1.Evaluate("x1<-rnorm(n1)");
            //o1 = sc1.GetSymbol("x1");

            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            string filepath = openFileDialog1.FileName;
            string extension = Path.GetExtension(filepath);
            string constr, sheetname;

            constr = string.Empty;

            switch (extension)
            {
                case ".xls":
                    constr = string.Format(Excel03conString, filepath, null);
                    break;

                case ".xlsx":
                    constr = string.Format(Excel07conString, filepath, null);
                    break;

                case ".csv":
                    constr = string.Format(Excel07conString, filepath, null);
                    break;
            }

            using (OleDbConnection cn = new OleDbConnection(constr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = cn;
                    cn.Open();
                    DataTable dt = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetname = dt.Rows[0]["Table_Name"].ToString();
                    cn.Close();
                }
            }

            using (OleDbConnection cn = new OleDbConnection(constr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        DataTable dt = new DataTable();
                        DataSet ds = new DataSet();
                        string query = @"select year as t,q1 from [" + sheetname + "]";
                        cmd.CommandText = query;
                        cmd.Connection = cn;
                        da.SelectCommand = cmd;
                        da.Fill(ds);
                        da.Fill(dt);
                        dataGridView1.DataSource = dt;
                        chart1.DataSource = ds;

                        DataView source = new DataView(ds.Tables[0]);
                        chart1.DataSource = source;
                        chart1.Series[0].YValueMembers = "q1";

                        chart1.Series[0].XValueMember = "t";
                        
                        chart1.DataBind();

                        //chart1.Series.Add("Year");
                        //chart1.Series.Add("Q1");
                        //chart1.Series["Year"].YValueMembers = "Year";
                        //chart1.Series["Q1"].XValueMember = "Q1";
                        //chart1.Series["Q3"].XValueMember = "Q3";
                        //chart1.Series["Q4"].XValueMember = "Q4";
                        //chart1.Series["Q5"].XValueMember = "Q5";
                        //chart1.Series["Y"].YValueMembers = "Y";
                        //chart1.Series["N"].YValueMembers = "N";
                        cn.Close();
                    }
                }
            }
        }
    }
}
