using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Task1
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SmartController controller = new SmartController(); ;
            controller.ReadExcelFile();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            SmartController controller = new SmartController();
            int retcode = controller.ReadExcelFile();
            
            if (retcode != 0)
            {
                MessageBox.Show(controller.GetLastException());
            }

            dataGridView1.DataSource = controller.GetDataTable();

        }
    }
}
