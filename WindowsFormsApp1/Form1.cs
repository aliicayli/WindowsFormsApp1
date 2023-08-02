using Aspose.Cells.GridDesktop;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();



            openFileDialog.Title = "Exel Dosyası Seçin";

            DialogResult result = openFileDialog.ShowDialog();

            if (result == DialogResult.OK)

            {

                //Seçilen dosyanın yolu

                string filePath = openFileDialog.FileName;

                //Seçilen dosyayı işlemek için kodlarınız
                try
                {
                    Process.Start("soffice.exe", filePath);
                }
                catch (System.ComponentModel.Win32Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();



            openFileDialog.Title = "Exel Dosyası Seçin";

            DialogResult result = openFileDialog.ShowDialog();

            if (result == DialogResult.OK)

            {

                //Seçilen dosyanın yolu

                string filePath = openFileDialog.FileName;

                //Seçilen dosyayı işlemek için kodlarınız
                try
                {
                    Form2 form2 = new Form2();
                    form2.Show();
                    form2.gridDesktop1.ImportExcelFile(filePath);

                    // Dosyayı GridDesktop kontrolüne yükle
                }
                catch (System.ComponentModel.Win32Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                }


            }
        }
    }
}
