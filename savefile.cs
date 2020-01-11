using System;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace _1task
{
    public static class savefile
    {
        public static void Savefile(DataSet a)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Execl files (*.xlsx)|*.xls";
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkBook = excelApp.Workbooks.Add(Type.Missing);

                foreach (DataTable table in a.Tables)
                {
                    Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                    excelWorkSheet.Name = table.TableName;

                    for (int j = 0; j < table.Rows.Count; j++)
                    {
                        for (int k = 0; k < table.Columns.Count; k++)
                        {
                            excelWorkSheet.Cells[j + 1, k + 1] = table.Rows[j].ItemArray[k].ToString();
                        }
                    }
                }
                MessageBox.Show("Вы успешно сохранили файл " + saveFileDialog1.FileName);
                excelWorkBook.SaveAs(saveFileDialog1.FileName);
                excelWorkBook.Close();
                excelApp.Quit();
            }
        }
    }
}
