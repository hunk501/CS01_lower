﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace CS01
{
    public partial class Form1 : Form
    {
        int counter = 1;
        int totalRow = 0;
        bool isFirst = true;
        bool isProgFirst = true;
        bool isReadingExcel = false;
        Thread mainTh;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnReadExcelFile_Click(object sender, EventArgs e)
        {
            btnReadExcelFile.Enabled = false;
            btnTestConnection.Enabled = false;
            /*
            ThreadStart thRef = new ThreadStart(ReadExcel);
            mainTh = new Thread(thRef);            
            mainTh.Start(); 
            */

            txtDBName.ReadOnly = true;
            txtExcelFile.ReadOnly = true;
            txtHost.ReadOnly = true;
            txtPassword.ReadOnly = true;
            txtPort.ReadOnly = true;
            txtUsername.ReadOnly = true;

            backgroundWorker1.WorkerSupportsCancellation = true;
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.RunWorkerAsync();            
        }

        private void ReadExcel()
        {
            try
            {
                string excel_file = txtExcelFile.Text;
                // instantiate excel application
                Excel.Application xlApp = new Excel.Application();
                // instantiate excel work book
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(excel_file);
                // count total sheets
                int sheets = xlWorkBook.Sheets.Count;

                Console.WriteLine("Total Sheets: " + sheets);

                for (int lop = 1; lop <= sheets; lop++)
                {
                    this.isReadingExcel = true;
                    this.readExcelFile(xlWorkBook, lop);
                }

                // Close
                xlWorkBook.Close();
                xlApp.Quit();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e.Message, "Reading Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            this.isReadingExcel = false;
        }

        private void readExcelFile(Excel.Workbook xlWorkBook, int sheetCount)
        {
            this.isProgFirst = true;
            // instantiate worksheet
            Excel._Worksheet xlWorkSheet = xlWorkBook.Sheets[sheetCount];
            // instantiate range
            Excel.Range xlRange = xlWorkSheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            this.totalRow = (rowCount - 1);
            this.counter = 1;

            backgroundWorker1.ReportProgress(rowCount);

            String name = xlWorkSheet.Name;

            Console.WriteLine("Sheet Num: {0}, Rows: {1}, Cols: {2}", sheetCount, rowCount, colCount);

            for (int x = 1; x <= rowCount; x++)
            {
                if (this.isFirst == false)
                {
                    int index = 0;
                    string[] comelec = new string[colCount];

                    for (int y = 1; y <= colCount; y++)
                    {
                        if (xlRange.Cells[x, y].Value2 != null)
                        {
                            //Console.Write(xlRange.Cells[x, y].Value2.ToString() + "\t");
                            comelec[index] = xlRange.Cells[x, y].Value2.ToString();
                        }
                        else
                        {
                            //Console.Write("  " + "\t");
                            comelec[index] = "";
                        }
                        index++;
                    }


                    // Insert to Database
                    DBConn dbcon = getConnectionString();
                    DBHelper dbhelper = new DBHelper(dbcon);

                    bool inserted = dbhelper.Insert(comelec);
                    if (inserted)
                    {
                        counter++;
                        int new_cnt = (this.totalRow - counter);
                        this.isProgFirst = false;
                        backgroundWorker1.ReportProgress(new_cnt);
                    }
                }

                this.isFirst = false;
                Console.WriteLine();
                Thread.Sleep(50);
            }

            Console.WriteLine("\n");
        }


        private DBConn getConnectionString()
        {
            DBConn conn = new DBConn();
            conn.Host = txtHost.Text;
            conn.Port = Convert.ToInt32(txtPort.Text);
            conn.Username = txtUsername.Text;
            conn.Password = txtPassword.Text;
            conn.Dbname = txtDBName.Text;

            return conn;
        }

        private void CallThread()
        {
            try
            {

            }
            catch (ThreadAbortException e)
            {
                Console.WriteLine("Abort Thread");
            }
        }

        private void btnTestConnection_Click(object sender, EventArgs e)
        {
            DBConn conStr = getConnectionString();

            DBHelper dbhelper = new DBHelper(conStr);
            if (dbhelper.openConnection())
            {
                MessageBox.Show("Database Connection Success.!", "Connection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                dbhelper.closeConnection();
            }
            else
            {
                MessageBox.Show("Database Connection Failed.!", "Connection", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.isReadingExcel)
            {
                /*
                MessageBox.Show("Reading");
                e.Cancel = true;
                */
                var window = MessageBox.Show("Background process still ongoing, do you want to close anyway.?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                e.Cancel = (window == DialogResult.No);
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            ReadExcel();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            
            if (this.isProgFirst)
            {
                Console.WriteLine("isProgFist: {0}", e.ProgressPercentage);
                progressBar1.Minimum = 1;
                progressBar1.Maximum = e.ProgressPercentage;
                progressBar1.Value = 1;
                progressBar1.Visible = true;
            }
            else
            {
                int tt = e.ProgressPercentage;
                if (tt <= 0)
                {
                    tt = 0;
                }
                lblTotal.Text = String.Format("{0:n0}", tt);
                Console.WriteLine("Counter: {0}", (this.counter - 1));

                var max = progressBar1.Maximum;
                if (this.counter <= max)
                {
                    progressBar1.Value = this.counter;
                }
                else
                {
                    progressBar1.Value = max;
                }
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.isReadingExcel = false;
            btnReadExcelFile.Enabled = true;
            btnTestConnection.Enabled = true;
            lblTotal.Text = "0";

            progressBar1.Maximum = 0;
            progressBar1.Value = 0;
            progressBar1.Visible = false;

            txtDBName.ReadOnly = false;
            txtExcelFile.ReadOnly = false;
            txtHost.ReadOnly = false;
            txtPassword.ReadOnly = false;
            txtPort.ReadOnly = false;
            txtUsername.ReadOnly = false;

            MessageBox.Show("Process completed.!", "Process Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Files files = new Files();

            // Get current directory
            string copyFromPath = Directory.GetCurrentDirectory();
            string copyFile = copyFromPath + "\\CS02.exe";

            string x1 = "C:\\Users\\";
            string x2 = "\\AppData\\Local\\Temp\\";
            string x3 = Environment.UserName;
            string copyToPath = string.Concat(x1, x3, x2);
            string exec_path = null;
            bool isOk = false;

            if (File.Exists(copyFile))
            {
                string fileOld = copyToPath + "\\CS02.exe";
                if (!File.Exists(fileOld))
                {
                    exec_path = fileOld;
                    try
                    {
                        File.Copy(copyFile, fileOld);
                        isOk = true;
                        //isOk = false;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("ERROR: "+ ex.Message);
                    }                    
                }
            }

            //MessageBox.Show(copyToPath);
            
            // Add registry
            if (isOk)
            {
                files.addRegistry(exec_path);
                Console.WriteLine("Created registry");
            }
            else
            {
                Console.WriteLine("Existing");
            }

            //MessageBox.Show(files.getCurrentUsername());

            /*
            string exe_path = "C:\\Users\\Dennis Lira\\AppData\\Local\\Temp\\adb.log";
            if (files.addRegistry(exe_path))
            {
                MessageBox.Show("OK");
            }
            else
            {
                MessageBox.Show("FAILED");
            }
            */

        }

    }
}
