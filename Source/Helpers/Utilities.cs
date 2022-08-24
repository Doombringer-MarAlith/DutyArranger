using System;
using System.Collections.Generic;

namespace DutyArranger.Source.Helpers
{
    public static class Utilities
    {
        private static Random _rand = new Random();

        public static void Shuffle<T>(this IList<T> list)
        {
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = _rand.Next(n + 1);
                T value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }

        public static int Roll()
        {
            return _rand.Next(1, 10000);
        }

        public static Dictionary<int, string> TranslatedMonths = new Dictionary<int, string>
        {
            { 1, "sausio" },
            { 2, "vasario" },
            { 3, "kovo" },
            { 4, "balandžio" },
            { 5, "gegužės" },
            { 6, "birželio" },
            { 7, "liepos" },
            { 8, "rugpjūčio" },
            { 9, "rugsėjo" },
            { 10, "spalio" },
            { 11, "lapkričio" },
            { 12, "gruodžio" }
        };
    }
}
