using omrapplication.Subjective;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace omrapplication
{
    public partial class StartScreen : Form
    {
        public StartScreen()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 obj = new Form1();
            obj.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SubSheetProcessor obj = new SubSheetProcessor();
            obj.ShowDialog();
        }
    }
}
