using System;
using System.Collections.Generic;

namespace AESTest2._0.Extensions
{
    public static class ListExtensions
    {
        public static void Shuffle<T>(this IList<T> l, int count)
        {
            Random rnd = new Random();
            int n = 0;
            if (count >= l.Count)
            {
                count = l.Count;
            }
            while (n < count)
            {
                int k = rnd.Next(n, count);
                T value = l[k];
                l[k] = l[n];
                l[n] = value;
                n++;
            }
        }
    }
}
