using CsharpExcel.helper;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CsharpExcel
{
    public partial class FrmReader : Form
    {
        public FrmReader()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            string rootpath = Path.Combine(System.Environment.CurrentDirectory, "read");
            var files = Directory.GetFiles(rootpath, "*.xls");
            float sumOfFloder = 0;//当前文件夹下所有充电
            foreach (var item in files)
            {
                FileInfo fi = new FileInfo(item);
                var dt=NpoiHelper.ImportExcelFile(fi.FullName);//通过NPOI读取excel
                //dgv.DataSource = dt;//测试时候, 看下内容是否正确
                //break;

                var sum = GetSum(dt, "充电量");// dt.Compute("sum(充电量)", "");//由于读取时候全部作为字符串处理,所以这个方法不正确的
                sumOfFloder += sum;
                txtResult.AppendText(fi.Name+"\t"+sum.ToString()+Environment.NewLine);
            }
            txtResult.AppendText("充电量合计:" + "\t" + sumOfFloder.ToString() + Environment.NewLine);
        }

        private float GetSum(DataTable dt, string columnName)//DataTable某一列
        {
            float sum = 0;
            foreach (DataRow dr in dt.Rows)
            {
                float tmp = 0;
                float.TryParse(dr[columnName].ToString(),out tmp);
                sum += tmp;
            }
            return sum;
        }






        /// <summary>
        /// 2016.5.11因为阿桦发给我的excel文件名前面有一个全角的空格,所以需要处理后才能正常排序.,用Trim()是无效的.
        /// </summary>
        /// <param name="fileName">输入的文件名</param>
        private void ReNameFiles(string fileName)
        {
            var newname = fileName;
            if (fileName[0].ToString() != "佛")
            {
                newname = fileName.Substring(1);
            }
            var source = Path.Combine(System.Environment.CurrentDirectory, "read", fileName);//拼接目录
            var dest = Path.Combine(System.Environment.CurrentDirectory, "read", newname);//连接目录
            File.Move(source, dest);
        }
    }
}
