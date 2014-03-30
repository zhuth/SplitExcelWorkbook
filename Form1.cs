using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;

namespace SplitExcelWorkbook
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            label1.Text = trackBar1.Value.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog();
            if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
            Workbook w = new Workbook();
            w.Open(ofd.FileName);

            string fileprefix = ofd.FileName;
            fileprefix = fileprefix.Substring(0, fileprefix.LastIndexOf('.')) + "_part";

            progressBar1.Style = ProgressBarStyle.Marquee;
            button1.Enabled = false;

            Worksheet s = w.Worksheets[0];
            int cols = 26;
            for (; s.Cells[0, cols].Value != null; cols++) ;

            int parts = trackBar1.Value;
            Workbook[] ws = new Workbook[parts];
            for (int i = 0; i < parts; ++i)
            {
                ws[i] = new Workbook();
                for (int j = 0; j < cols; ++j)
                {
                    ws[i].Worksheets[0].Cells[0, j].PutValue(s.Cells[0, j].Value);
                }
            }

            int maxRows = (int)numericUpDown1.Value;
            if (maxRows == 0) maxRows = int.MaxValue;

            for (int i = 1; s.Cells[i, 0].Value != null && i <= maxRows; ++i)
            {
                for (int j = 0; j < cols; ++j)
                    ws[i % parts].Worksheets[0].Cells[((i % parts == 0) ? 0 : 1) + i / parts, j].PutValue(s.Cells[i, j].Value);

                Application.DoEvents();
            }

            for (int i = 0; i < parts; ++i)
            {
                ws[i].Save(fileprefix + i + ".xls");

                Application.DoEvents();
            }
            progressBar1.Style = ProgressBarStyle.Blocks;
            button1.Enabled = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
            var sfd = new SaveFileDialog();
            if (sfd.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            var bs = new Workbook[ofd.FileNames.Length];
            for (int i = 0; i < bs.Length; ++i)
            {
                bs[i] = new Workbook(); bs[i].Open(ofd.FileNames[i]);
            }
            var b = new Workbook();

            int rows = 0, cols = 26;
            // copy the first workbook
            for (; bs[0].Worksheets[0].Cells[rows, cols].Value != null; ++cols) ;

            for (; bs[0].Worksheets[0].Cells[rows, 0].Value != null; ++rows)
                for (int j = 0; j < cols; ++j)
                    b.Worksheets[0].Cells[rows, j].PutValue(bs[0].Worksheets[0].Cells[rows, j].Value);

            for (int bi = 1; bi < bs.Length; ++bi)
            {
                for (int ri = 1; bs[bi].Worksheets[0].Cells[ri, 0].Value != null; ++ri, ++rows)
                    for (int j = 0; j < cols; ++j)
                        b.Worksheets[0].Cells[rows, j].PutValue(bs[bi].Worksheets[0].Cells[ri, j].Value);
            }

            b.Save(sfd.FileName);
        }
    }
}
