using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace HybridPr
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var marvelHeroes = new List<MarvelHero>
            {
                new MarvelHero
                {
                    SuperheroName = "Железный человек",
                    RealName = "Тони Старк",
                    Superpower = "Интеллект, сверхкостюм",
                    Origin = "Земля",
                    Team = "Мстители",
                    DateOfBirth = new DateTime(1970, 5, 29),
                    DateOfDeath = null,
                    Actor = "Роберт Даун Мл."
                },
                new MarvelHero
                {
                    SuperheroName = "Капитан Америка",
                    RealName = "Стив Роджер",
                    Superpower = "Сверхчеловек",
                    Origin = "Земля",
                    Team = "Мстители",
                    DateOfBirth = new DateTime(1918, 7, 4),
                    DateOfDeath = new DateTime(2023, 5, 4),
                    Actor = "Крис Эванс"
                }
            };

            Console.WriteLine("Хотите добавить нового персонажа? (да/нет)");
            string response = Console.ReadLine().ToLower();

            if (response == "да")
            {
                var newHero = new MarvelHero();

                Console.Write("Имя супергероя: ");
                newHero.SuperheroName = Console.ReadLine();

                Console.Write("Настоящее имя: ");
                newHero.RealName = Console.ReadLine();

                Console.Write("Суперсила: ");
                newHero.Superpower = Console.ReadLine();

                Console.Write("Происхождение: ");
                newHero.Origin = Console.ReadLine();

                Console.Write("Команда: ");
                newHero.Team = Console.ReadLine();

                Console.Write("Дата рождения (гггг-мм-дд): ");
                newHero.DateOfBirth = DateTime.Parse(Console.ReadLine());

                Console.Write("Дата смерти (гггг-мм-дд, оставьте пустым если жив): ");
                string deathDate = Console.ReadLine();
                if (!string.IsNullOrWhiteSpace(deathDate))
                {
                    newHero.DateOfDeath = DateTime.Parse(deathDate);
                }

                Console.Write("Актер: ");
                newHero.Actor = Console.ReadLine();

                marvelHeroes.Add(newHero);
            }

            var thisAddIn = new ThisAddIn();
            thisAddIn.ThisAddIn_Startup(marvelHeroes);
        }
    }
}
