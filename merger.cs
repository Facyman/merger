using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;

namespace _1task
{
    class merger
    {
        public static void Merger(bool cont,int Levenstein)
        {           
            bool match = false;
            int neighbour = 0;
            OpenFileDialog Dialog = new OpenFileDialog();
            Dialog.Multiselect = true;
            DialogResult result = Dialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                DataSet newsheets = new DataSet();
                foreach (String filename1 in Dialog.FileNames) // весь цикл используется 1ин раз для добавления первого файла в пустой файл
                {
                    foreach (String filename2 in Dialog.FileNames)
                    {
                        if (filename1 != filename2) // не проверять 2 одинаковых файла
                        {
                            try
                            {
                                cont = true;
                                string Location1 = filename1;
                                string SourceConstr1 = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Location1 + "';Extended Properties= 'Excel 12.0;HDR=No;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text;'";
                                OleDbConnection Connection1 = new OleDbConnection(SourceConstr1);
                                string Location2 = filename2;
                                string SourceConstr2 = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Location2 + "';Extended Properties= 'Excel 12.0;HDR=No;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text;'";
                                OleDbConnection Connection2 = new OleDbConnection(SourceConstr2);

                                Connection1.Open();
                                Connection2.Open();

                                DataTable datatable1 = new DataTable();
                                datatable1 = Connection1.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                                String[] excelSheets1 = new String[datatable1.Rows.Count];
                                int i = 0;

                                foreach (DataRow row in datatable1.Rows)
                                {
                                    excelSheets1[i] = row["TABLE_NAME"].ToString();
                                    i++;
                                }

                                DataTable datatable2 = new DataTable();
                                datatable2 = Connection2.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                                String[] excelSheets2 = new String[datatable2.Rows.Count];
                                int j = 0;

                                foreach (DataRow row in datatable2.Rows)
                                {
                                    excelSheets2[j] = row["TABLE_NAME"].ToString();
                                    j++;
                                }

                                foreach (string addsheet1 in excelSheets1) // список листов 1 файла
                                {
                                    foreach (string addsheet2 in excelSheets2) // список листов 2 файла
                                    {
                                        OleDbDataAdapter dataAdapter1 = new OleDbDataAdapter("SELECT * FROM [" + addsheet1 + "]", Connection1);
                                        DataTable table1 = new DataTable();
                                        dataAdapter1.Fill(table1);

                                        OleDbDataAdapter dataAdapter2 = new OleDbDataAdapter("SELECT * FROM [" + addsheet2 + "]", Connection2);
                                        DataTable table2 = new DataTable();
                                        dataAdapter2.Fill(table2);

                                        int rowCount1 = table1.Rows.Count;
                                        int colCount1 = table1.Columns.Count;
                                        int rowCount2 = table2.Rows.Count;
                                        int colCount2 = table2.Columns.Count;

                                        if (match == false && tools.LevenshteinDistance(addsheet1, addsheet2) <= Levenstein && tools.identicalColimns(table1, table2) == true) // addsheet1 == addsheet2 длина левенштайна
                                        {
                                            DataTable newtable = newsheets.Tables.Add(addsheet1.Remove(addsheet1.Length - 1));
                                            for (int k = 0; k < colCount1; k++)
                                            {
                                                for (int l = 0; l < colCount2; l++)
                                                {
                                                    if (table1.Rows[0][k].ToString() == table2.Rows[0][l].ToString())
                                                    {
                                                        if (newtable.Rows.Count < table1.Rows.Count)
                                                        {
                                                            dataAdapter1.Fill(newtable);
                                                        }
                                                        match = true;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                Connection1.Close();
                                Connection1.Dispose();
                                Connection2.Close();
                                Connection2.Dispose();
                                datatable1.Dispose();
                                datatable2.Dispose();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }
                    }
                }

                if (match == false)
                    MessageBox.Show("Выберите файлы которые возможно соединить");

                if (cont == true && match == true) // нашлось более 2ух файлов и у любых 2ух файлов есть хотя бы 1ин общий лист => выполнение главной части программы
                {
                    for (int z = 0; z < Dialog.FileNames.Count(); z++) // повторение цикла для полного копирования всех данных
                    {
                        foreach (String filename in Dialog.FileNames)
                        {
                            string Location = filename;
                            string SourceConstr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Location + "';Extended Properties= 'Excel 12.0;HDR=No;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text;'";
                            using (OleDbConnection Connection = new OleDbConnection(SourceConstr))
                            {
                                Connection.Open();
                                DataTable datatable = new DataTable();
                                datatable = Connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                String[] excelSheets = new String[datatable.Rows.Count];
                                int w = 0;
                                foreach (DataRow row in datatable.Rows)
                                {
                                    excelSheets[w] = row["TABLE_NAME"].ToString();
                                    w++;
                                }

                                String[] excelSheets2 = new String[newsheets.Tables.Count];
                                for (int s = 0; s < newsheets.Tables.Count; s++)
                                {
                                    excelSheets2[s] = newsheets.Tables[s].ToString();
                                }

                                foreach (string addsheet in excelSheets) // список листов  файла откуда берем данные
                                {
                                    foreach (string newsheet in excelSheets2) // список листов конечного файла
                                    {
                                        string addsheetx = addsheet.Remove(addsheet.Length - 1);//удаление $ так как datarow возвращает значение с $

                                        OleDbDataAdapter dataAdapter1 = new OleDbDataAdapter("SELECT * FROM [" + addsheet + "]", Connection);
                                        DataTable addtable = new DataTable();
                                        dataAdapter1.Fill(addtable);
                                        DataTable newtable = newsheets.Tables[newsheet.ToString()];

                                        int rowCountadd = addtable.Rows.Count;
                                        int colCountadd = addtable.Columns.Count;

                                        int colCountnew = newtable.Columns.Count;


                                        if (tools.samesheetswithlevenstein(addsheetx, excelSheets2, Levenstein) && tools.identicalColimns(addtable, newtable) == true)  // проверка если есть хотя бы 1 одинаковый столб 
                                        {
                                            if (tools.LevenshteinDistance(addsheetx, newsheet) <= Levenstein) // (addsheetx == newsheet) длина левенштайна
                                            {
                                                string[] champ = new string[1000]; // хранит номера строк которые нужно исключить чтобы остались только номера строк которые нужно добавить как новые
                                                newtable = tools.notidenticalColimns(addtable, newtable); // добавление несуществующих колонок
                                                for (int addr = 1; addr < rowCountadd; addr++)
                                                {

                                                    bool permission = true;
                                                    int ovo = 0;
                                                    int rowCountnew = newtable.Rows.Count;
                                                    for (int newr = 1; newr < rowCountnew; newr++)
                                                    {
                                                        neighbour = 0; // переменные для определения добавить новую строку либо дополнить строку                                                            
                                                        for (int i = 0; i < colCountadd; i++)
                                                        {
                                                            for (int j = 0; j < colCountnew; j++)
                                                            {
                                                                if (addtable.Rows[0][i].ToString() == newtable.Rows[0][j].ToString()) // выбор только тех колонн которые одинаковы по названию
                                                                {
                                                                    if (addtable.Rows[addr][i].ToString() == newtable.Rows[newr][j].ToString() | (newtable.Rows[newr][j] == DBNull.Value && addtable.Rows[addr][i] != DBNull.Value)) // если информация в ячейках 2ух таблиц идентична 
                                                                    {                                                                                                                                                                // или одна из ячеек пуста                                                                          
                                                                        neighbour++;
                                                                        permission = false;
                                                                    }
                                                                }
                                                            }
                                                        }  //НИЖНИЙ ЦИКЛ ДЛЯ ДОБАВЛЕНИЯ ДАННЫХ В ПОЧТИ ИДЕНТИЧ РЯДЫ    
                                                        for (int i = 0; i < colCountadd; i++)
                                                        {
                                                            for (int j = 0; j < colCountnew; j++)
                                                            {
                                                                if (addtable.Rows[0][i].ToString() == newtable.Rows[0][j].ToString()) //поиск соседей . одинак столбы
                                                                {
                                                                    if (newtable.Rows[newr][j] == DBNull.Value && neighbour == colCountadd)
                                                                    {

                                                                        newtable.Rows[newr][j] = addtable.Rows[addr][i]; // добавление ячейки в этом ряду

                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }


                                                    if (permission == false)
                                                    {
                                                        champ[ovo] = addr.ToString();
                                                        ovo++;
                                                    }



                                                    if (!champ.Contains(addr.ToString())) // ЦИКЛ ДЛЯ ДОБАВЛЕНИЯ НОВОГО РЯДА  
                                                    {
                                                        newtable.Rows.Add();
                                                        for (int i = 0; i < colCountadd; i++)
                                                        {
                                                            for (int j = 0; j < colCountnew; j++)
                                                            {
                                                                if (addtable.Rows[0][i].ToString() == newtable.Rows[0][j].ToString())
                                                                {
                                                                    newtable.Rows[rowCountnew][j] = addtable.Rows[addr][i]; //добавление нового ряда                                                     
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (!tools.samesheetswithlevenstein(addsheetx, excelSheets2, Levenstein))
                                            {
                                                DataTable oldtable = newsheets.Tables.Add(addsheetx);
                                                OleDbDataAdapter newadapter = new OleDbDataAdapter("SELECT * FROM [" + addsheet + "]", Connection);
                                                newadapter.Fill(oldtable);
                                            }
                                        }
                                    }

                                }
                                Connection.Close();
                                Connection.Dispose();
                            }
                        }
                    }
                }
                else
                    MessageBox.Show("Выберите более 1ного файла");
            }
        }
    }
}
