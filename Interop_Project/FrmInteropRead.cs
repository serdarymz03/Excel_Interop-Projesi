using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Interop_Project
{
    public partial class FrmInteropRead : Form
    {
        public FrmInteropRead()
        {
            InitializeComponent();
        }

        private void FrmInteropRead_Load(object sender, EventArgs e)
        {

        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Dosyası |*.xlsx";

            DialogResult dialogResult = openFileDialog1.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                string fullPath = openFileDialog1.FileName;
                ReadExcel(fullPath);
            }
        }

        void ReadExcel(string fullPath)
        {
            Microsoft.Office.Interop.Excel.Application application;
            Workbook workbook;
            Worksheet worksheet;

            application = new Microsoft.Office.Interop.Excel.Application();
            workbook = application.Workbooks.Open(fullPath);
            worksheet = workbook.Worksheets[1];

            Range range = worksheet.UsedRange;
            //range = worksheet.Range["C6", "F15"];

            List<Person> people = new List<Person>();

            int colIndex = 1;
            for (int j = 2; j <= range.Rows.Count; j++)
            {
                Person person = new Person(
                    (int)range.Cells[j, colIndex].Value,
                    range.Cells[j, ++colIndex].Value.ToString(),
                    (DateTime)range.Cells[j, ++colIndex].Value
                    );
                people.Add(person);
                colIndex = 1;
            }

            application.Quit();
            DtgPersonList.DataSource = people;
        }
    }

}
