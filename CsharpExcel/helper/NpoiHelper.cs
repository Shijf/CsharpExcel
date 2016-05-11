using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsharpExcel.helper
{
    public class NpoiHelper
    {

        /// <summary>
        /// 通过文件名读取excel的内容. 第一行为标题, 内容从第二行开始
        /// </summary>
        /// <param name="filePath">要读取的excel文件完整路径</param>
        /// <returns>返回DataTable</returns>
        public static DataTable ImportExcelFile(string filePath)
        {
            HSSFWorkbook hssfworkbook;
            #region//初始化信息
            try
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    hssfworkbook = new HSSFWorkbook(file);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            #endregion

            NPOI.SS.UserModel.ISheet sheet = hssfworkbook.GetSheetAt(0);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            DataTable dt = new DataTable();
            for (int j = 0; j < (sheet.GetRow(0).LastCellNum); j++)
            {
                var columnName = sheet.GetRow(0).GetCell(j).ToString();
                dt.Columns.Add(columnName);
            }
            rows.MoveNext();//跳过第一行, 第一行为标题
            while (rows.MoveNext())
            {
                HSSFRow row = (HSSFRow)rows.Current;
                DataRow dr = dt.NewRow();
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    NPOI.SS.UserModel.ICell cell = row.GetCell(i);
                    if (cell == null)
                        dr[i] = null;
                    else
                        dr[i] = cell.ToString();
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }



    }
}
