using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Linq;
using DataTable = System.Data.DataTable;

namespace Nöbetci
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("---- GÜNÜN NÖBETÇİSİ ------\n");

            ReadExceltoDataTable("Ocak", @"C:\Users\senev\Documents\VS\Nöbet.xlsx", 1, 1);

            Console.ReadLine();
        }

        public static DataTable ReadExceltoDataTable(string worksheetName, string saveLocation, int headerLine, int ColumnStart)
        {
            DataTable dt = new DataTable();
            Application xlApp;
            Workbook xlWorkbook;
            Worksheet xlWorksheet;
            Range xlRange;

            try
            {
                xlApp = new Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                xlWorkbook = xlApp.Workbooks.Open(saveLocation);
                xlWorksheet = (Worksheet)xlApp.Worksheets.Item[worksheetName];
                xlRange = xlWorksheet.UsedRange;

                int columsCount = xlRange.Columns.Count;
                int rowCount = xlRange.Rows.Count;

                for (int j = ColumnStart; j <= columsCount; j++)
                {
                    dt.Columns.Add(Convert.ToString(xlRange.Cells[headerLine, j].value), typeof(string));
                }

                for (int i = headerLine + 1; i <= rowCount; i++)
                {
                    DataRow dr = dt.NewRow();
                    for (int j = ColumnStart; j <= columsCount; j++)
                    {
                        dr[j - ColumnStart] = Convert.ToString(xlRange.Cells[i, j].value);
                    }

                    dt.Rows.InsertAt(dr, dt.Rows.Count + 1);
                }

                NobetciOlustur(dt);

                xlWorkbook.Close();
                xlApp.Quit();
                return dt;


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }

        public static void NobetciOlustur(DataTable _dt)
        {
            DateTime today = DateTime.Today;
            string dToday = today.ToString("D");

            DataColumn[] columns = _dt.Columns.Cast<DataColumn>().ToArray();

            bool nobetciCheck = _dt.AsEnumerable().Any(row => columns.Any(col => row[col].ToString() == dToday));

            if (nobetciCheck)
            {
                var dateRow = _dt.AsEnumerable().Select((r, i) => new { Row = r, Index = i })
                    .Where(x => (string)x.Row["Tarih"] == dToday.ToString()).FirstOrDefault();

                int rowNumber = 0;

                if (dateRow != null)
                {
                    rowNumber = dateRow.Index;

                    for (int i = 0; i < _dt.Columns.Count - 4; i++)
                    {
                        GununNobetcisi(_dt, rowNumber, i);
                    }

                    SmsGonder(_dt, rowNumber);
                }
            }
            else
            {
                Console.WriteLine("Bugün Nöbetçi Yok");
            }
        }
        
        public static void GununNobetcisi(DataTable _dt, int _rowNumber, int _colNumber)
        {
            string nobetciBilgileri = _dt.Rows[_rowNumber].ItemArray[_colNumber].ToString();

            Console.WriteLine("" + nobetciBilgileri);
        }

        public static void SmsGonder(DataTable _dt, int _rowNumber)
        {
            Console.WriteLine("\nSayın " + _dt.Rows[_rowNumber].ItemArray[0] + ", bugün nobetçisiniz.");
        } 
    }
}
