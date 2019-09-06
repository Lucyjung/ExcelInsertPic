using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Core;
using excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;

namespace ExcelInsertPic
{
    public partial class Excel : Form
    {
        public Excel()
        {
            InitializeComponent();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            List<string> filelist = new List<string>();
            foreach (string file in listBox1.Items)
            {
                filelist.Add(file);
            }
            BtnWriteSpreedSheetClick(textBox1.Text, filelist.ToArray());
        }
        private void BtnWriteSpreedSheetClick(String outputPath, String[] fileList)
        {
            var xlApp = new excel.Application();
            String outputFile = @outputPath + @"\picture.xls";
            excel.Workbook xlWorkBook = xlApp.Workbooks.Add();
            excel.Worksheet xlWorkSheet = xlWorkBook.Sheets[1];
            xlWorkSheet.PageSetup.PaperSize = excel.XlPaperSize.xlPaperA4;

            int INIT_X = Int32.Parse(textStartX.Text);
            int INIT_Y = Int32.Parse(textStartY.Text);
            int startX = INIT_X;
            int staryY = INIT_Y;

            int pictureSizeX = Int32.Parse(width.Text);
            int pictureSizeY = Int32.Parse(height.Text);
            int gap = Int32.Parse(pageGap.Text);
            int i = 0;
            foreach (String file in fileList)
            {
                var shape = xlWorkSheet.Shapes.AddPicture(file, MsoTriState.msoFalse, MsoTriState.msoCTrue, startX, staryY, pictureSizeX, pictureSizeY);
                shape.Line.Style = MsoLineStyle.msoLineSingle;
                int row = (i / 2) + 1;
                if (i % 2 == 0)
                {
                    startX += pictureSizeX;
                }
                else
                {
                    staryY += pictureSizeY;
                    startX = INIT_X;
                }
                i++;
                if ((i %8 ) == 0)
                {
                    staryY += gap;
                }
            }
            

            xlWorkBook.SaveAs(outputFile, excel.XlFileFormat.xlWorkbookNormal);
            xlWorkBook.Close(true);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("File " + outputFile + " is created !");
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Images (*.BMP;*.JPG;*.GIF;*.PNG)|*.BMP;*.JPG;*.GIF;*.PNG|" +
                        "All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (String file in openFileDialog.FileNames)
                    {
                        listBox1.Items.Add(file);
                    }
                }
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    textBox1.Text = fbd.SelectedPath;
                }
            }
        }
        private void number_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }
    }
}
