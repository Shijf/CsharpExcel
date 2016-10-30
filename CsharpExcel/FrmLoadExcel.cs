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
    public partial class FrmLoadExcel : Form
    {
        public FrmLoadExcel()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            var filename = "902.xls";
            var dt = NpoiHelper.ImportExcelFile(filename);//通过NPOI读取excel
            dgv.DataSource = dt;
        }

        List<DateTime> times = new List<DateTime>();
        private void button1_Click(object sender, EventArgs e)
        {
            times.Clear();
            var maxcol = dgv.Columns.Count ;
            foreach (DataGridViewRow dr in dgv.Rows)
            {
                for (int i = 1; i < maxcol; i++)
                {
                    var cell=dr.Cells[i].Value.ToString();
                    Console.Write(cell + " ");
                    if (cell != "")
                    {
                        DateTime time = DateTime.Parse(cell);
                        times.Add(time);
                    }
                }
                Console.WriteLine();
            }

            int count = 0;
            foreach (var item in times)
            {
                count++;
                Console.WriteLine(count.ToString()+" "+item.ToString());
            }
        }
    }
}
