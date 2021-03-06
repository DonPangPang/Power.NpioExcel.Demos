using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TestApp1
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        public ExcelStructType StructType { get; set; } = ExcelStructType.DropList;

        public string DropList { get; set; } = "";

        public string Name { get; set; } = null!;
        public int Index { get; set; } = 0;

        public int Width { get; set; } = 10;

        // public ExcelColumnAttribute(string name, string index) { Name = name; Index = index; }
    }

    //[AttributeUsage(AttributeTargets.Property)]
    //public class ExcelColumnAttribute : Attribute
    //{
    //    public ExcelStructType StructType { get; set; } = ExcelStructType.String;

    // public string Name { get; set; } = null!; public int Index { get; set; } = 0;

    // public int Width { get; set; } = 10;

    //    // public ExcelColumnAttribute(string name, string index) { Name = name; Index = index; }
    //}
}