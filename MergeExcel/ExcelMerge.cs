using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
// using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MergeExcel
{
    public partial class ExcelMerge : Form
    {
        #region private variable define

        private string szleft;

        private string szright;

        private string szOut;

        private DataTable dt;

        #endregion

        public ExcelMerge()
        {
            InitializeComponent();
        }                

        private void btnEmrge_Click(object sender, EventArgs e)
        {
            Merge mergeFile = new Merge(szleft, szright);
            mergeFile.MergeEx();
            dt = mergeFile.GetDataTable();       
            this.ultraGrid1.DataSource = dt;
            this.ultraGrid1.Refresh();
            this.ultraGrid1.Show();
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "(*.xls)|*.xls";
            dlg.RestoreDirectory = true;
            dlg.ShowDialog();
            cbxLeft.Items.Add(dlg.FileName);
            szleft = dlg.FileName;
            cbxLeft.Items.Add(szleft);
            cbxLeft.Text = szleft;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "(*.xls)|*.xls";
            dlg.RestoreDirectory = true;
            dlg.ShowDialog();
            cbxRight.Items.Add(dlg.FileName);
            szright = dlg.FileName;
            cbxRight.Items.Add(szright);
            cbxRight.Text = szright;
        }       

        private void cbxLeft_SelectedIndexChanged(object sender, EventArgs e)
        {
            szleft = cbxLeft.SelectedItem.ToString();
        }

        private void cbxRight_SelectedIndexChanged(object sender, EventArgs e)
        {
            szright = cbxRight.SelectedItem.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExportGrid exportGrid = new ExportGrid(this.ultraGrid1, szOut);
            exportGrid.ExportToFile();
        }
        
    }
}
