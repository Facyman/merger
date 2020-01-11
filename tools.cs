using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace _1task
{
    public static class tools
    {
        public static bool identicalColimns(DataTable addtable, DataTable newtable) // ПОИСК СТОЛБОВ С ОДИНАКОВЫМИ НАЗВАНИЯМИ
        {

            int colCountadd = addtable.Columns.Count;
            int colCountnew = newtable.Columns.Count;
            bool check = false;
            for (int i = 0; i < colCountadd; i++)
            {
                for (int j = 0; j < colCountnew; j++)
                {
                    if (addtable.Rows[0][i].ToString() == newtable.Rows[0][j].ToString())
                    {
                        check = true;
                        break;
                    }
                }
                if (check == true)
                {
                    break;
                }
            }
            return check;
        }
        public static DataTable notidenticalColimns(DataTable addtable, DataTable newtable) // ПОИСК НЕСУЩЕСТВУЮЩИХ СТОЛБОВ
        {
            int colCountadd = addtable.Columns.Count;
            int colCountnew = newtable.Columns.Count;
            int g = 0;
            for (int i = 0; i < colCountadd; i++)
            {
                g = 0;
                for (int j = 0; j < colCountnew; j++)
                {
                    if (addtable.Rows[0][i].ToString() == newtable.Rows[0][j].ToString())
                    {
                        g++;
                    }
                    if (g == 0 && j == colCountnew - 1)
                    {
                        newtable.Columns.Add();
                        newtable.Rows[0][colCountnew] = addtable.Rows[0][i];
                        colCountnew++;
                    }

                }
            }
            return newtable;
        }

        public static bool samesheetswithlevenstein(string addsheet, string[] newsheets, int levenstein) // СУЩЕСТВУЕТ ЛИ УЖЕ ТАКОЙ СТОЛБ В ОБЩЕЙ ТАБЛИЦЕ(ЛЕВЕНШТАЙН)
        {
            bool exist = false;
            int newsheetscount = newsheets.Count();
            for (int i = 0; i < newsheetscount; i++)
            {
                if (LevenshteinDistance(addsheet, newsheets[i]) <= levenstein)
                    exist = true;
            }
            return exist;
        }
        public static int LevenshteinDistance(string firstWord, string secondWord) // ДЛИНА ЛЕВЕНШТАЙНА
        {
            var n = firstWord.Length + 1;
            var m = secondWord.Length + 1;
            var matrixD = new int[n, m];

            const int deletionCost = 1;
            const int insertionCost = 1;

            for (var i = 0; i < n; i++)
            {
                matrixD[i, 0] = i;
            }

            for (var j = 0; j < m; j++)
            {
                matrixD[0, j] = j;
            }

            for (var i = 1; i < n; i++)
            {
                for (var j = 1; j < m; j++)
                {
                    var substitutionCost = firstWord[i - 1] == secondWord[j - 1] ? 0 : 1;

                    matrixD[i, j] = Math.Min(matrixD[i - 1, j] + deletionCost,
                                     Math.Min(matrixD[i, j - 1] + insertionCost, matrixD[i - 1, j - 1] + substitutionCost));
                }
            }

            return matrixD[n - 1, m - 1];
        }
    }
}
