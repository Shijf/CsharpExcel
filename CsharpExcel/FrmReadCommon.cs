using CsharpExcel.helper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CsharpExcel
{
    public partial class FrmReadCommon : Form
    {
        public FrmReadCommon()
        {
            InitializeComponent();
            this.Text = "读取一般的excel文件";
        }

        private void FrmReadCommon_Load(object sender, EventArgs e)
        {

        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = Application.StartupPath;
            var dr=ofd.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                var file = ofd.SafeFileName;
               // MessageBox.Show(file);
                LoadFile(file);
            }
        }

        private void LoadFile(string file)
        {
            var dt=NpoiHelper.ImportExcelFile(file);
            dgv.DataSource = dt;
        }
    }
}
