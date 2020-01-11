using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
namespace _1task
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataSet newsheets = new DataSet();
        bool cont = false;
    
        private void button1_Click(object sender, EventArgs e)
        {
            newsheets.Clear();
            {
                int Levenstein = (int)levenstein.zero;
                if (radioButton1.Checked == true)
                {
                    Levenstein = (int)levenstein.zero;
                }
                if (radioButton2.Checked == true)
                {
                    Levenstein = (int)levenstein.one;
                }
                if (radioButton3.Checked == true)
                {
                    Levenstein = (int)levenstein.two;
                }
                merger.Merger(cont, Levenstein);
            }
        }      
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                radioButton1.Checked = false;
                radioButton3.Checked = false;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                radioButton2.Checked = false;
                radioButton3.Checked = false;
            }
           
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked == true)
            {
                radioButton1.Checked = false;
                radioButton2.Checked = false;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (cont == true)
            {
                savefile.Savefile(newsheets);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}

