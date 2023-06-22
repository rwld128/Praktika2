using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Praktika2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            tableLayoutPanel3.Visible = false;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                tableLayoutPanel3.Visible = true;
                tableLayoutPanel2.Visible = false;
            }
            else
            {
                tableLayoutPanel2.Visible = true;
                tableLayoutPanel3.Visible = false;
            }
        }
    }
}
