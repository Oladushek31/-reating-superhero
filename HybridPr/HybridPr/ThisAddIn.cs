using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace HybridPr
{
    internal class ThisAddIn
    {
        public void ThisAddIn_Startup(List<MarvelHero> marvelHeroes)
        {
            DisplayInExcel(marvelHeroes, (hero, cell) =>
            {
                cell.Value = hero.SuperheroName;
                cell.Offset[0, 1].Value = hero.RealName;
                cell.Offset[0, 2].Value = hero.Superpower;
                cell.Offset[0, 3].Value = hero.Origin;
                cell.Offset[0, 4].Value = hero.Team;
                cell.Offset[0, 5].Value = hero.DateOfBirth.ToShortDateString();
                cell.Offset[0, 6].Value = hero.DateOfDeath.HasValue ? hero.DateOfDeath.Value.ToShortDateString() : "";
                cell.Offset[0, 7].Value = hero.Actor;
            });
        }

        void DisplayInExcel(IEnumerable<MarvelHero> heroes, Action<MarvelHero, Excel.Range> DisplayFunc)
        {
            var excelApp = new Excel.Application();
            excelApp.Workbooks.Add();
            excelApp.Visible = true;

            var headerLabels = new string[] { "Геройское имя", "Настоящее имя", "Суперспособность", "Происхождение", "Команда", "Дата рождения", "Дата смерти", "Актер" };
            for (int i = 0; i < headerLabels.Length; i++)
            {
                excelApp.Cells[1, i + 1].Value = headerLabels[i];
            }

            int row = 2;
            foreach (var hero in heroes)
            {
                DisplayFunc(hero, excelApp.Cells[row, 1]);
                row++;
            }

            excelApp.Columns.AutoFit();
        }
    }
}
