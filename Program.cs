// See https://aka.ms/new-console-template for more information
using System.Diagnostics;
using TestApp1;

// var data = "123465"; var guid = "84D7478D-310D-40BA-886B-EAE61977EA9E"; var datetime =
// "2022/05/25"; var enumd = 1; var bool_false = "False"; var bool_true = "True"; var bool_num = 1;
// var data_byte = data.ConvertTo<byte>(0); var data_int = data.ConvertTo<int>(0); var data_long =
// data.ConvertTo<long>(0); var data_decimal = data.ConvertTo<decimal>(0); var data_float =
// data.ConvertTo<float>(0); var data_double = data.ConvertTo<double>(0); var data_enum =
// enumd.ConvertTo<TestEnum>(TestEnum.Default); var data_dateTime =
// datetime.ConvertTo<DateTime>(DateTime.MinValue); var data_guid =
// guid.ConvertTo<Guid>(Guid.Empty); var data_bool = bool_false.ConvertTo<bool>(); var data_bool2 =
// bool_true.ConvertTo<bool>(); var data_bool3 = bool_num.ConvertTo<bool>();

// Console.WriteLine($"{data_byte} {data_byte.GetType().Name}"); Console.WriteLine($"{data_int}
// {data_int.GetType().Name}"); Console.WriteLine($"{data_long} {data_long.GetType().Name}");
// Console.WriteLine($"{data_decimal} {data_decimal.GetType().Name}");
// Console.WriteLine($"{data_float} {data_float.GetType().Name}"); Console.WriteLine($"{data_double}
// {data_double.GetType().Name}"); Console.WriteLine($"{data_enum} {data_enum.GetType().Name}");
// Console.WriteLine($"{data_dateTime} {data_dateTime.GetType().Name}");
// Console.WriteLine($"{data_guid} {data_guid.GetType().Name}"); Console.WriteLine($"{data_bool}
// {data_bool.GetType().Name}"); Console.WriteLine($"{data_bool2} {data_bool2.GetType().Name}");
// Console.WriteLine($"{data_bool3} {data_bool3.GetType().Name}");

var rootPath = "E:/Demos/Test/TestApp1/Power.NpioExcel.Demos";
//var filePath = "E:/Demos/Test/TestApp1/20220527025949.xlsx";
//var excel = new ExcelImport(filePath);
//var data = excel.ToList<User>();

//Console.WriteLine($"姓名\t年龄\t性别\t生日");
//foreach (var item in data)
//{
//    Console.WriteLine($"{item.Name}\t{item.Age}\t{item.Gender}\t{item.Born}");
//}
//Console.WriteLine($"Error Lines : {string.Join(",", excel.Errors)}");

var outPath = $"{rootPath}/{DateTime.Now.ToString("yyyyMMddhhmmss")}.xlsx";

//var outData = new List<OutUser>
//{
//    // new OutUser
//    // {
//    //     Name = "A",
//    //     Age = 1,
//    //     Gender = "男",
//    //     Born = DateTime.Now
//    // },
//    // new OutUser
//    // {
//    //     Name = "B",
//    //     Age = 1,
//    //     Gender = "男",
//    //     Born = DateTime.Now
//    // },
//    // new OutUser
//    // {
//    //     Name = "C",
//    //     Age = 1,
//    //     Gender = "男",
//    //     Born = DateTime.Now
//    // },
//    // new OutUser
//    // {
//    //     Name = "D",
//    //     Age = 1,
//    //     Gender = "男",
//    //     Born = DateTime.Now
//    // },
//};

//new ExcelExport<OutUser>().Export(outData).ToFile(outPath);

//Stopwatch stopwatch = new Stopwatch();

//var filePath = "E:/Demos/Test/TestApp1/Power.NpioExcel.Demos/20220602030849.xlsx"; // 65535条数据
////var excel = new ExcelImport(filePath);
////var data = excel.ToList<User>();

//stopwatch.Start();
//var excel1 = new ExcelImport(filePath);
//foreach (var data in excel1.ToListAwait<User>())
//{
//    //// data 为 500 条数据
//    //Console.WriteLine($"姓名\t年龄\t性别\t生日\t 条数:{data.Count}");
//    //foreach (var item in data)
//    //{
//    //    Console.WriteLine($"{item.Name}\t{item.Age}\t{item.Gender}\t{item.Born}");
//    //}
//    //Console.WriteLine($"Error Lines : {string.Join(",", excel.Errors)}");
//}

//stopwatch.Stop();

//Console.WriteLine($"500条分批转换耗费时间: {stopwatch.Elapsed}");

//stopwatch.Restart();
//var excel2 = new ExcelImport(filePath);
//var dt = excel2.ToList<User>();

//stopwatch.Stop();

//Console.WriteLine($"直接转换耗费时间: {stopwatch.Elapsed}");

//Console.ReadKey();

new ExcelExport<OutUser>().Export(new List<OutUser>()).ToTemp(outPath);