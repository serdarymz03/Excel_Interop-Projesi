using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Konsol
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Interop_Read
            /*
            Application application;
            Workbook workbook;
            Worksheet worksheet;

            application = new Application();
            workbook = application.Workbooks.Open(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Excel Örnek\\Örnek_Veri.xlsx");
            worksheet = workbook.Worksheets[1];

            Range range = worksheet.UsedRange;
            //range = worksheet.Range["C6", "F15"];

            List<object> list = new List<object>();

            for (int i = 1; i <= range.Columns.Count; i++)
            {
                for (int j = 1; j < range.Rows.Count; j++)
                {
                    list.Add(range.Cells[j, i].Value);
                }
            }

            foreach (var item in list)
            {
                Console.WriteLine(item + " / ");
            }
            application.Quit();
            */
            #endregion

            #region Interop_Write
            /*
            List<Person> people = new List<Person>()
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

            Application application;
            Workbook workbook;
            Worksheet worksheet;

            application = new Application();
            application.Visible = true;
            workbook = application.Workbooks.Add("");

            worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "Id";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Cells[1, 3].Value = "BirthDate";
            worksheet.Range["A1", "C1"].Font.Bold = true;
            worksheet.Range["A1", "C1"].Font.Color = XlRgbColor.rgbDarkRed;
            worksheet.Range["A1", "C1"].Font.Size = 15;


            int rowIndex = 2, columnIndex = 1;
            foreach (Person item in people)
            {
                worksheet.Cells[rowIndex, columnIndex].Value = item.Id;
                worksheet.Cells[rowIndex, ++columnIndex].Value = item.Name;
                worksheet.Cells[rowIndex, ++columnIndex].Value = item.BirthDate;
                rowIndex++;
                columnIndex = 1;
            }
            worksheet.Range["A2", "C11"].Font.Size = 13;
            application.Visible = false;
            workbook.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Excel Örnek\\Yeni_Veri.xlsx");
            application.Quit();

            Console.WriteLine("İşlem Tamamlandı");
            */
            #endregion

            Console.ReadLine();
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
