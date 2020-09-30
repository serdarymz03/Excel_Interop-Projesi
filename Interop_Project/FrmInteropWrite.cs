using Microsoft.Office.Interop.Excel;
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

namespace Interop_Project
{
    public partial class FrmInteropWrite : Form
    {
        List<Person> people;
        string fullPath;
        public FrmInteropWrite()
        {
            InitializeComponent();
        }

        private void FrmInterop_Load(object sender, EventArgs e)
        {
            people = new List<Person>()
            {
                new Person(1,"Melda",new DateTime(1995,10,15)),
                new Person(2,"Kerem",new DateTime(1990,1,25)),
                new Person(3,"Tugay",new DateTime(1993,7,19)),
                new Person(4,"Ümit",new DateTime(1991,5,29)),
                new Person(5,"Damla",new DateTime(1999,9,17)),
                new Person(6,"Gizem",new DateTime(1990,11,23)),
                new Person(7,"Hülya",new DateTime(1989,12,13)),
                new Person(8,"Hüseyin",new DateTime(1988,7,26)),
                new Person(9,"Kemal",new DateTime(2000,1,14)),
                new Person(10,"Selin",new DateTime(1997,10,24))
            };
            DtgPersonList.DataSource = people;

            DtgPersonList.Columns["Id"].HeaderText = "No";
            DtgPersonList.Columns["Name"].HeaderText = "Ad Soyad";
            DtgPersonList.Columns["BirthDate"].HeaderText = "Doğum Tarihi";

        }

        void Export_Table(DataGridView dataGridView, string fullPath)
        {
            Microsoft.Office.Interop.Excel.Application application;
            Workbook workbook;
            Worksheet worksheet;

            application = new Microsoft.Office.Interop.Excel.Application();
            application.Visible = true;
            workbook = application.Workbooks.Add("");

            worksheet = workbook.ActiveSheet;

            int headerColumnIndex = 0;
            foreach (DataGridViewColumn item in dataGridView.Columns)
            {
                worksheet.Cells[1, ++headerColumnIndex].Value = item.HeaderText;
            }

            worksheet.Range["A1", "C1"].Font.Bold = true;
            worksheet.Range["A1", "C1"].Font.Color = XlRgbColor.rgbDarkRed;
            worksheet.Range["A1", "C1"].Font.Size = 15;


            int rowIndex = 2, columnIndex = 1;
            foreach (DataGridViewRow item in dataGridView.Rows)
            {
                worksheet.Cells[rowIndex, columnIndex].Value = item.Cells["Id"].Value;
                worksheet.Cells[rowIndex, ++columnIndex].Value = item.Cells["Name"].Value;
                worksheet.Cells[rowIndex, ++columnIndex].Value = item.Cells["BirthDate"].Value;
                rowIndex++;
                columnIndex = 1;
            }
            worksheet.Range["A2", "C11"].Font.Size = 13;
            application.Visible = false;
            workbook.SaveAs(fullPath);
            application.Quit();
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (DtgPersonList.DataSource == null)
            {
                return;
            }
            if (string.IsNullOrEmpty(fullPath))
            {
                MessageBox.Show("Dosya Yolu Seçiniz");
                return;
            }
            Export_Table(DtgPersonList, fullPath);
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            //saveFileDialog1.DefaultExt = "xlsx";
            //saveFileDialog1.AddExtension = true;
            DialogResult dialogResult = saveFileDialog1.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                fullPath = saveFileDialog1.FileName;

                string oldName = Path.GetFileName(fullPath);
                string newName = Path.GetFileNameWithoutExtension(oldName) + ".xlsx";

                fullPath = fullPath.Replace(oldName, newName);
                TxtPath.Text = fullPath;
            }
        }
    }

    class Person
    {
        int id; string name; DateTime birthDate;

        public int Id { get => id; set => id = value; }
        public string Name { get => name; set => name = value; }
        public DateTime BirthDate { get => birthDate; set => birthDate = value; }

        public Person(int id, string name, DateTime birthDate)
        {
            this.id = id;
            this.name = name;
            this.birthDate = birthDate;
        }
    }
}
