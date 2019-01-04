using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System;
using System.IO;
using System.Text;
namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            label2.Text = openFileDialog1.ToString();
            label3.Text = openFileDialog1.FileName.ToString();
            string path;
            path = openFileDialog1.FileName.ToString();
            textBox1.Text = File.ReadAllText(path);



        }

        private void button2_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
            textBox2.Text = saveFileDialog1.FileName.ToString();
            string path2;
            path2 = saveFileDialog1.FileName.ToString();
            File.WriteAllText(path2, textBox1.Text);

        }
    }
}
