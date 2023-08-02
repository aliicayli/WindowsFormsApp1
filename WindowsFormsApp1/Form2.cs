using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {


        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            var result = System.Windows.MessageBox.Show("Excel dosyasını kaydetmek istiyor musunuz?", "Kaydet", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

            // Dialogdan dönen sonuca göre işlem yapın
            switch (result)
            {
                case MessageBoxResult.Yes:
                    // Evet seçilirse, dosya seçici dialog oluşturun
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    saveFileDialog.Title = "Excel dosyasını kaydet";

                    DialogResult resultt = saveFileDialog.ShowDialog();

                    if (resultt == DialogResult.OK)
                    {
                        FileStream fs = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.ReadWrite);
                        gridDesktop1.ExportExcelFile(fs,Aspose.Cells.GridDesktop.FileFormatType.Excel2007Xlsx);
                        fs.Close();
                    }
                    break;
                case MessageBoxResult.No:
                    break;
            }
        }
    }
}
