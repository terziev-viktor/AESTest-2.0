using System;
using System.Collections.Generic;

namespace AESTest2._0.Extensions
{
    public static class ListExtensions
    {
        public static void Shuffle<T>(this IList<T> l)
        {
            Random rnd = new Random();
            int n = l.Count;
            while (n > 1)
            {
                n--;
                int k = rnd.Next(n + 1);
                T value = l[k];
                l[k] = l[n];
                l[n] = value;
            }
        }
    }
}
