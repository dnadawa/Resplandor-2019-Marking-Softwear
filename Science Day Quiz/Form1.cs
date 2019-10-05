using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Science_Day_Quiz
{
    public partial class Form1 : Form



    {

        public static int teamAmarks = 0;
        public static int teamBmarks = 0;
        int fixedMarks = 5;


        void giveMarks(String team)
        {
            if(team == "TeamA")
            {
                teamAmarks = teamAmarks + fixedMarks;
            }
            else if(team == "TeamB")
            {
                teamBmarks = teamBmarks + fixedMarks;
            }

            sendtoExcel();
        }

        void sendtoExcel() {
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open("C:\\Users\\DULAJ\\Desktop\\RCSS.xlsx");
                Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                x.Range["A2"].Value = teamAmarks.ToString();
                x.Range["B2"].Value = teamBmarks.ToString();
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




        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            giveMarks("TeamA");
            lblteama.Text = teamAmarks.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            giveMarks("TeamA");
            lblteama.Text = teamAmarks.ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            giveMarks("TeamA");
            lblteama.Text = teamAmarks.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            giveMarks("TeamA");
            lblteama.Text = teamAmarks.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            giveMarks("TeamB");
            lblteamb.Text = teamBmarks.ToString();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            giveMarks("TeamB");
            lblteamb.Text = teamBmarks.ToString();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            giveMarks("TeamB");
            lblteamb.Text = teamBmarks.ToString();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            giveMarks("TeamB");
            lblteamb.Text = teamBmarks.ToString();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            int marks = Int32.Parse(txtteama.Text);
            teamAmarks = marks;
            lblteama.Text = teamAmarks.ToString();
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open("C:\\Users\\DULAJ\\Desktop\\RCSS.xlsx");
                Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                x.Range["A2"].Value = marks.ToString();
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

        private void button10_Click(object sender, EventArgs e)
        {
            int marks = Int32.Parse(txtteamb.Text);
            teamBmarks = marks;
            lblteamb.Text = teamBmarks.ToString();
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open("C:\\Users\\DULAJ\\Desktop\\RCSS.xlsx");
                Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                x.Range["B2"].Value = marks.ToString();
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


