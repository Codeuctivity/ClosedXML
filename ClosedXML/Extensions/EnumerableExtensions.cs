
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal static class EnumerableExtensions
    {
        public static void ForEach<T>(this IEnumerable<T> source, Action<T> action)
        {
            foreach (var item in source)
            {
                action(item);
            }
        }

        public static Type GetItemType(this IEnumerable source)
        {
            return GetGenericArgument(source?.GetType());

            Type GetGenericArgument(Type collectionType)
            {
                if (collectionType == null)
                {
                    return null;
                }

                var ienumerable = collectionType.GetInterfaces()
                    .SingleOrDefault(i => i.GetGenericArguments().Length == 1 &&
                                          i.Name == "IEnumerable`1");

                return ienumerable?.GetGenericArguments()?.FirstOrDefault();
            }
        }

        public static bool HasDuplicates<T>(this IEnumerable<T> source)
        {
            var distinctItems = new HashSet<T>();
            foreach (var item in source)
            {
                if (!distinctItems.Add(item))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
