using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TestApp1
{
    public static class FormatExtensions
    {
        public static T? ConvertTo<T>(this object? value, T? defaultValue = default(T))
        {
            if (value is null)
                return defaultValue;

            try
            {
                return (T)Convert.ChangeType(value, typeof(T));
            }
            catch
            {
                if (typeof(T) == typeof(Guid))
                    return (T)Convert.ChangeType(Guid.Parse(value.ToString() ?? ""), typeof(T));

                if (typeof(T) == typeof(DateTime))
                    return (T)Convert.ChangeType(DateTime.Parse(value.ToString() ?? ""), typeof(T));

                if (typeof(T).IsEnum)
                    return (T)Convert.ChangeType(Enum.Parse(typeof(T), value.ToString() ?? ""), typeof(T));


                return defaultValue;
            }
        }
    }

    public enum TestEnum
    {
        Test1 = 0,
        Test2 = 1,
        Test3 = 2,
        Test4 = 3,

        Default = 999
    }
}