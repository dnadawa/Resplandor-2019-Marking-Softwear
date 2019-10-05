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
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 g = new Form1();
            g.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Round_2 g = new Round_2();
            g.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Round_3 g = new Round_3();
            g.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form1.teamAmarks = 0;
            Form1.teamBmarks = 0;
            Round_2.totAmarks = 0;
            Round_2.totBmarks = 0;

            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open("C:\\Users\\DULAJ\\Desktop\\RCSS.xlsx");
                Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                x.Range["A2"].Value = "0";
                x.Range["B2"].Value = "0";
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

        private void button4_Click(object sender, EventArgs e)
        {
            Backup g = new Backup();
            g.Show();
        }
    }
}
