using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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
        public OutUser()
        {
            var props = this.GetType().GetProperties();

            var t = props.Where(x => x.GetCustomAttribute<ExcelColumnAttribute>().Index == 0).FirstOrDefault();
            t.SetValue("DropList", "部门1,部门2,部门3");
        }

        [ExcelColumn(Name = "姓名", Index = 0)]
        public string Name { get; set; } = null!;

        [ExcelColumn(Name = "年龄", Index = 1)]
        public int Age { get; set; }

        [ExcelColumn(Name = "部门", Index = 2, StructType = ExcelStructType.DropList)]
        public string Department { get; set; } = null!;

        [ExcelColumn(Name = "性别", Index = 3)]
        public string Gender { get; set; } = null!;

        [ExcelColumn(Name = "生日", Index = 4, Width = 30)]
        public DateTime Born { get; set; }
    }
}