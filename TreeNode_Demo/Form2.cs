using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using RDotNet;
using System.IO;

namespace TreeNode_Demo
{
    public partial class Form2 : Form  
    {
        public Form2()
        {
            InitializeComponent();
            //string dlldir = @"C:/Program Files/R/R-2.13.0/bin/x64";
            //string path = Path.GetFullPath(dlldir).Replace(@"/", @"/");
            //REngine.SetEnvironmentVariables(path);
            //REngine engine = REngine.GetInstance("R");
            string dllDir = @"C:/Program Files/R/R-2.13.0/bin/i386";
            string path = Path.GetFullPath(dllDir).Replace(@"/", @"/");

            //    //code solution from http://stackoverflow.com/questions/7960738/importing-mgcv-fails-because-rlapack-dll-cannot-be-found

            //    string rhome = System.Environment.GetEnvironmentVariable("R_HOME");
            //    if (string.IsNullOrEmpty(rhome))
            //        rhome = @"C:\Program Files\R\R-2.13.0\bin\i386";

            //     System.Environment.SetEnvironmentVariable("R_HOME", rhome);
            //     System.Environment.SetEnvironmentVariable("PATH", System.Environment.GetEnvironmentVariable("PATH") + ";" + rhome + @"bini386");
            //    // Set the folder in which R.dll locates.
            ////REngine.SetDllDirectory(@”C:Program FilesRR-2.12.0bini386″);
            //REngine.SetDllDirectory(@"C:Program FilesRR-2.15.0bini386");
            //using (REngine engine = REngine.CreateInstance("RDotNet", new[] { "-q" }))  // quiet mode
            //{
            //    engine.e
            //}
            REngine.SetEnvironmentVariables(path);
            //REngine.GetInstance("RDotNet");
            REngine engine = REngine.GetInstance("R");
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //IStatConnector test1 = new StatConnector();
            //test1.Init("R");
                        //test1.Evaluate("x<-rnorm(20)");
            //test1.EvaluateNoReturn("hist(x)"); 



            //DataTable dtb = new DataTable();
            //dtb.Columns.Add("Column1", Type.GetType("System.String"));
            //dtb.Columns.Add("Column2", Type.GetType("System.String"));
            //DataRow dtr1 = dtb.NewRow();
            //dtr1[0] = "abc";
            //dtr1[1] = "cdf";
            //dtb.Rows.Add(dtr1);
            //DataRow dtr2 = dtb.NewRow();
            //dtr2[0] = "asdasd";
            //dtr2[1] = "cdasdasf";
            //dtb.Rows.Add(dtr2);

            //using (var engine = REngine.GetInstance("R"))
            //{
            //    string[,] stringData = new string[dtb.Rows.Count, dtb.Columns.Count];
            //    for (int row = 0; row < dtb.Rows.Count; row++)
            //    {
            //        for (int col = 0; col < dtb.Columns.Count; col++)
            //        {
            //            stringData[row, col] = dtb.Rows[row].ItemArray[col].ToString();
            //        }
            //    }
            //    CharacterMatrix matrix = engine.CreateCharacterMatrix(stringData);
            //    engine.SetSymbol("myRDataFrame", matrix);
            //    engine.Evaluate("myRDataFrame <- as.data.frame(myRDataFrame, stringsAsFactors = FALSE)");
            //    engine.Evaluate("str(myRDataFrame)");

            //}
            //Console.ReadKey();


            //string dlldir = @"C:\Program Files\R\R-2.13.0\bin\x64";
            //string path = Path.GetFullPath(dlldir).Replace(@"\", @"\");
            ////REngine.;
            //REngine.SetEnvironmentVariables(path);

            REngine engine = REngine.GetInstance("R");
            try
            {
                // import csv file
                engine.Evaluate("dataset<-read.table(file.choose(), header=TRUE, sep = ',')");

                // retrieve the data frame
                DataFrame dataset = engine.Evaluate("dataset").AsDataFrame();

                if (dataset.RowCount > 1 && dataset.ColumnCount > 1)
                {
                    // 06-05
                    //# get means for variables in data frame mydata
                    //# excluding missing values 
                    //engine.Evaluate("sapply(dataset, mean, na.rm=TRUE)");
                    //engine.Evaluate("attach(mtcars)");
                    //engine.Evaluate("plot(wt, mpg)");
                    //engine.Evaluate("abline(lm(mpg~wt))");
                    //engine.Evaluate("title('Regression of MPG on Weight')");

                    
                    // -- 05- 05
                    //var x = engine.CreateNumericVector(Enumerable.Range(0, 13).Select(i => i * Math.PI / 12).ToArray());
                    //var cos = engine.GetSymbol("cos").AsFunction();
                    //var y = cos.Invoke(new[] { x }).AsNumeric();
                    ////Console.WriteLine(string.Join(" ", y));

                    // 06-05
                    //engine.Evaluate("plot(dataset)"); 

                    for (int i = 0; i < dataset.ColumnCount; ++i)
                    {
                        dataGridView1.ColumnCount++;
                        dataGridView1.Columns[i].Name = dataset.ColumnNames[i];
                    }
                    for (int i = 0; i < dataset.RowCount; ++i)
                    {
                        dataGridView1.RowCount++;
                        dataGridView1.Rows[i].HeaderCell.Value = dataset.RowNames[i];
                        for (int k = 0; k < dataset.ColumnCount; ++k)
                        {
                            dataGridView1[k, i].Value = dataset[i, k];
                        }
                    }

                    //engine.Evaluate("plot(pie(table(dataset$Year)))"); 
                    //other info
                    //engine.Evaluate("colors = c('red', 'yellow', 'green', 'violet','orange', 'blue', 'pink", 'cyan')");
                    engine.Evaluate("plot(pie(table(dataset)))");

                    //# Creating a Graph

                }
                else {
                    MessageBox.Show("OOPS..! Data Not Found. Please Select Excel or .CSV File.");
                }
            }
            catch(Exception ex)
            {
                if (ex.Message == "Error in file.choose() : file choice cancelled\n")
                {
                    MessageBox.Show(@"Please Select Excel or .CSV File."); 
                }
                else
                {
                    MessageBox.Show(@"Equation error."); 
                }
            }
        }
    }
}
