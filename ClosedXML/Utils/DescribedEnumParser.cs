﻿
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Utils
{
    internal static class DescribedEnumParser<T>
    {
        private static readonly Lazy<IDictionary<string, T>> fromDescriptions = new Lazy<IDictionary<string, T>>(() =>
        {
            return ParseEnumDescriptions().ToDictionary(a => a.Item2, a => a.Item1);
        });

        private static readonly Lazy<IDictionary<T, string>> toDescriptions = new Lazy<IDictionary<T, string>>(() =>
        {
            return ParseEnumDescriptions().ToDictionary(a => a.Item1, a => a.Item2);
        });

        public static T FromDescription(string value)
        {
            return fromDescriptions.Value[value];
        }

        public static bool IsValidDescription(string value)
        {
            return fromDescriptions.Value.ContainsKey(value);
        }

        public static string ToDescription(T value)
        {
            return toDescriptions.Value[value];
        }

        private static IEnumerable<Tuple<T, string>> ParseEnumDescriptions()
        {
            var type = typeof(T);
            return type.GetFields(BindingFlags.Public | BindingFlags.Static)
                .Select(f =>
                {
                    var attributes = f.GetCustomAttributes(typeof(DescriptionAttribute), inherit: false).OfType<DescriptionAttribute>();
                    var description = attributes.FirstOrDefault()?.Description ?? f.Name;
                    return new Tuple<T, string>
                    (
                        (T)Enum.Parse(type, f.Name),
                        description
                    );
                });
        }
    }
}
