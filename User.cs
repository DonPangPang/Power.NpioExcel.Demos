using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TestApp1
{
    public class User : IExcelStruct
    {
        [ExcelColumn(Name = "姓名", Index = 0)]
        public string Name { get; set; } = null!;

        [ExcelColumn(Name = "年龄", Index = 2)]
        public int Age { get; set; }

        [ExcelColumn(Name = "性别", Index = 3)]
        public string Gender { get; set; } = null!;

        [ExcelColumn(Name = "生日", Index = 1, Width = 30)]
        public DateTime Born { get; set; }

        [ExcelColumn(Name = "姓名1", Index = 0)]
        public string Name1 { get; set; } = null!;

        [ExcelColumn(Name = "年龄1", Index = 2)]
        public int Age1 { get; set; }

        [ExcelColumn(Name = "性别1", Index = 3)]
        public string Gender1 { get; set; } = null!;

        [ExcelColumn(Name = "生日1", Index = 1, Width = 30)]
        public DateTime Born1 { get; set; }

        [ExcelColumn(Name = "姓名2", Index = 0)]
        public string Name2 { get; set; } = null!;

        [ExcelColumn(Name = "年龄2", Index = 2)]
        public int Age2 { get; set; }

        [ExcelColumn(Name = "性别2", Index = 3)]
        public string Gender2 { get; set; } = null!;

        [ExcelColumn(Name = "生日2", Index = 1, Width = 30)]
        public DateTime Born2 { get; set; }
    }

    public class OutUser : IExcelStruct
    {
        [ExcelColumn(Name = "姓名", Index = 0)]
        public string Name { get; set; } = null!;

        [ExcelColumn(Name = "年龄", Index = 1)]
        public int Age { get; set; }

        [ExcelColumn(Name = "性别", Index = 2)]
        public string Gender { get; set; } = null!;

        [ExcelColumn(Name = "生日", Index = 3, Width = 30)]
        public DateTime Born { get; set; }
    }
}