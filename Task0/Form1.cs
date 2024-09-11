using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Task0
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();


            kA.KeyPress += new KeyPressEventHandler(NumericTextBox_KeyPress);
            kb.KeyPress += new KeyPressEventHandler(NumericTextBox_KeyPress);
            kY1.KeyPress += new KeyPressEventHandler(NumericTextBox_KeyPress);
            kY2.KeyPress += new KeyPressEventHandler(NumericTextBox_KeyPress);
        }

        string filePath;
        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            SmartController controller = new SmartController();
            int retcode = controller.ReadExcelFile();

            if (retcode != 0)
            {
                MessageBox.Show(controller.GetLastException());
            }
            else
            {
                filePath = controller.FilePath;
                dataGridView1.DataSource = controller.GetDataTable();
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (filePath != null)
            {
                ExcelExporter.ExportDataGridViewToExcel(dataGridView1, filePath);
            }
        }


        private void NumericTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Проверка, является ли нажатая клавиша цифрой или управляющим символом (например, Backspace)
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != ',' && e.KeyChar != '-')
            {
                e.Handled = true; // Отменить ввод символа
            }
        }

        private void kA_TextChanged(object sender, EventArgs e)
        {
            Clac();
        }

        void Clac()
        {
            try
            {
                double a = double.Parse(kA.Text);
                double y1 = double.Parse(kY1.Text);
                double y2 = double.Parse(kY2.Text);

                if (a == 0)
                {
                    return;
                }

                bool isSwaped = false;
                if (y1 > y2)
                {
                    double temp = y1;
                    y1 = y2;
                    y2 = temp;
                    isSwaped = true;
                }
                double xAverage = (y2 -y1) / (2 * a);

                //if (a < 0)
                //{
                //    xAverage *= -1;
                //}

                kXn.Text = xAverage.ToString();

                try
                {
                    double b = double.Parse(kb.Text);

                    double x1 = (y1 - b) / a;
                    double x2 = (y2 - b) / a;


                    
                    kXAverage.Text = (x1 + xAverage).ToString();
                    

                    if (isSwaped)
                    {
                        double temp = x1;
                        x1 = x2;
                        x2 = temp;
                    }

                    kX1.Text = x1.ToString();
                    kX2.Text = x2.ToString();



                }
                catch
                {

                }
            }
            catch
            {

            }
        }

        private void kY1_TextChanged(object sender, EventArgs e)
        {
            Clac();
        }

        private void kY2_TextChanged(object sender, EventArgs e)
        {
            Clac();
        }

        private void kb_TextChanged(object sender, EventArgs e)
        {
            Clac();
        }

        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
