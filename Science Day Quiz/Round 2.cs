using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Science_Day_Quiz
{
    public partial class Round_2 : Form

    {







        public Round_2()
        {
            InitializeComponent();
        }

        public static int teamAmarks = 0;
        public static int teamBmarks = 0;
        int preA;
        int preB;
        public static int totAmarks;
        public static int totBmarks;

        private void Round_2_Load(object sender, EventArgs e)
        {
             preA = Form1.teamAmarks;
             preB = Form1.teamBmarks;
            lblprea.Text = preA.ToString();
            lblpreb.Text = preB.ToString();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            teamAmarks = Int32.Parse(txtteama.Text);
            lblteama.Text = teamAmarks.ToString();
            totAmarks = preA + teamAmarks;
            tota.Text = totAmarks.ToString();
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open("C:\\Users\\DULAJ\\Desktop\\RCSS.xlsx");
                Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                x.Range["A2"].Value = totAmarks.ToString();
               
                //x.Range["C2"].Value = "N/A";
                //x.Range["D2"].Value = "N/A";
                //x.Range["c2"].Value = "how are you";
                //x.Cells[7, 2] = "Hello";
                sheet.Close(true, Type.Missing, Type.Missing);
                excel.Quit();
                MessageBox.Show("Success");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            teamBmarks = Int32.Parse(txtteamb.Text);
            lblteamb.Text = teamBmarks.ToString();
            totBmarks = preB + teamBmarks;
            totb.Text = totBmarks.ToString();
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open("C:\\Users\\DULAJ\\Desktop\\RCSS.xlsx");
                Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
              
                x.Range["B2"].Value = totBmarks.ToString();
                //x.Range["C2"].Value = "N/A";
                //x.Range["D2"].Value = "N/A";
                //x.Range["c2"].Value = "how are you";
                //x.Cells[7, 2] = "Hello";
                sheet.Close(true, Type.Missing, Type.Missing);
                excel.Quit();
                MessageBox.Show("Success");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            int marks = Int32.Parse(txtteamabackup.Text);
            teamAmarks = marks;
            totAmarks = preA + teamAmarks;
            tota.Text = totAmarks.ToString();
            lblteama.Text = teamAmarks.ToString();
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open("C:\\Users\\DULAJ\\Desktop\\RCSS.xlsx");
                Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                x.Range["A2"].Value = totAmarks.ToString();
                //x.Range["B2"].Value = teamBmarks.ToString();
                //x.Range["C2"].Value = "N/A";
                //x.Range["D2"].Value = "N/A";
                //x.Range["c2"].Value = "how are you";
                //x.Cells[7, 2] = "Hello";
                sheet.Close(true, Type.Missing, Type.Missing);
                excel.Quit();
                MessageBox.Show("Success");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int marks = Int32.Parse(txtteambbackup.Text);
            teamBmarks = marks;
            totBmarks = preB + teamBmarks;
            totb.Text = totBmarks.ToString();
            lblteamb.Text = teamBmarks.ToString();
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open("C:\\Users\\DULAJ\\Desktop\\RCSS.xlsx");
                Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                x.Range["B2"].Value = totBmarks.ToString();
                //x.Range["B2"].Value = teamBmarks.ToString();
                //x.Range["C2"].Value = "N/A";
                //x.Range["D2"].Value = "N/A";
                //x.Range["c2"].Value = "how are you";
                //x.Cells[7, 2] = "Hello";
                sheet.Close(true, Type.Missing, Type.Missing);
                excel.Quit();
                MessageBox.Show("Success");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex);
            }
        }
    }
}
ooh yeah motharfakers!!!
{get set for some fucking experience (true, type.missing, typemissing);
