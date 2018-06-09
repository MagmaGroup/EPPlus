using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;


namespace Magma.EPPlus.Tester
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (var ep = new ExcelPackage(true, new FileInfo(@"D:\Cloud\Magma Group LTD\Support Team - MES Projects\IAI\Issues.xlsx")))
            {
                var worksheet =  ep.Workbook.Worksheets.First();
                for (int i = 2; i<= worksheet.Dimension.Rows; i++)
                {
                    int? id = null;
                    if (worksheet.Cells[i, "id"].Value != null)
                    {
                        id = Convert.ToInt32(worksheet.Cells[i, "id"].Value);
                    }



                    //worksheet.Protection.
                    //worksheet.Cells[i, "title"].Style.Locked = true;
                    //worksheet.Protection
                    //Value.ToString();
                }

            }
        }
    }
}
