using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using NPOI.SS.UserModel;

#nullable disable

namespace TestApp1
{
    public class ExcelImport
    {
        public string FilePath { get; set; }
        public string SheetName { get; set; }

        public List<int> Errors { get; set; }

        private DataTable _dataTable;

        /// <summary>
        /// 初始化文件路径和Sheet名称
        /// </summary>
        /// <param name="filePath">  </param>
        /// <param name="sheetName"> </param>
        public ExcelImport(string filePath, string sheetName = "Sheet1")
        {
            FilePath = filePath;
            SheetName = sheetName;

            Errors = new List<int>();
        }

        /// <summary>
        /// 转为指定类型
        /// </summary>
        /// <typeparam name="T"> </typeparam>
        /// <returns> </returns>
        public List<T> ToList<T>() where T : IExcelStruct
        {
            if (_dataTable is null) this.ToDataTable();

            var list = new List<T>();
            var type = typeof(T);
            var properties = type.GetProperties();
            foreach (DataRow row in _dataTable.Rows)
            {
                var t = Activator.CreateInstance<T>();
                foreach (var property in properties)
                {
                    var prop_name = property.GetCustomAttribute<ExcelColumnAttribute>().Name;
                    var value = row[prop_name];
                    if (value != DBNull.Value)
                    {
                        var d = Convert.ChangeType(value, property.PropertyType);
                        property.SetValue(t, d);
                    }
                }
                list.Add(t);
            }
            return list;
        }

        public List<T> ToList<T>(DataTable dt) where T : IExcelStruct
        {
            var list = new List<T>();
            var type = typeof(T);
            var properties = type.GetProperties();
            foreach (DataRow row in dt.Rows)
            {
                var t = Activator.CreateInstance<T>();
                foreach (var property in properties)
                {
                    var prop_name = property.GetCustomAttribute<ExcelColumnAttribute>().Name;
                    var value = row[prop_name];
                    if (value != DBNull.Value)
                    {
                        var d = Convert.ChangeType(value, property.PropertyType);
                        property.SetValue(t, d);
                    }
                }
                list.Add(t);
            }
            return list;
        }

        public IEnumerable<List<T>> ToListAwait<T>() where T : IExcelStruct
        {
            foreach (var item in ToDataTableAwait())
            {
                yield return ToList<T>(item);
            }
        }

        /// <summary>
        /// Excel转为DataTable
        /// </summary>
        /// <returns> </returns>
        public DataTable ToDateTable()
        {
            return ToDateTable();
        }

        /// <summary>
        /// Excel转为DataTable
        /// </summary>
        /// <returns> </returns>
        public ExcelImport ToDataTable()
        {
            _dataTable = new DataTable();
            using (var fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read))
            {
                var workbook = WorkbookFactory.Create(fs);
                var sheet = workbook.GetSheet(SheetName);
                if (sheet == null)
                {
                    throw new Exception("Excel sheet not found");
                }

                var headerRow = sheet.GetRow(0);
                var headerRowCount = headerRow.LastCellNum;
                for (var i = headerRow.FirstCellNum; i < headerRowCount; i++)
                {
                    var column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                    _dataTable.Columns.Add(column);
                }

                var rowCount = sheet.LastRowNum;
                for (var i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    var isSave = true;

                    try
                    {
                        var row = sheet.GetRow(i);
                        var dataRow = _dataTable.NewRow();
                        for (var j = row.FirstCellNum; j < headerRowCount; j++)
                        {
                            if (row.GetCell(j) != null && row.GetCell(j).CellType != CellType.Blank)
                            {
                                dataRow[j] = row.GetCell(j).ToString();
                            }
                            else
                            {
                                isSave = false;
                                Errors.Add(i + 1);
                                break;
                            }
                        }
                        if (isSave) _dataTable.Rows.Add(dataRow);
                    }
                    catch
                    {
                        Errors.Add(i + 1);
                    }
                }
            }

            return this;
        }

        public IEnumerable<DataTable> ToDataTableAwait()
        {
            var dt = new DataTable();
            using (var fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read))
            {
                var workbook = WorkbookFactory.Create(fs);
                var sheet = workbook.GetSheet(SheetName);
                if (sheet == null)
                {
                    throw new Exception("Excel sheet not found");
                }

                var headerRow = sheet.GetRow(0);
                var headerRowCount = headerRow.LastCellNum;
                for (var i = headerRow.FirstCellNum; i < headerRowCount; i++)
                {
                    var column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                    dt.Columns.Add(column);
                }

                var rowCount = sheet.LastRowNum;
                for (var i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    var isSave = true;

                    try
                    {
                        var row = sheet.GetRow(i);
                        var dataRow = dt.NewRow();
                        for (var j = row.FirstCellNum; j < headerRowCount; j++)
                        {
                            if (row.GetCell(j) != null && row.GetCell(j).CellType != CellType.Blank)
                            {
                                dataRow[j] = row.GetCell(j).ToString();
                            }
                            else
                            {
                                isSave = false;
                                Errors.Add(i + 1);
                                break;
                            }
                        }
                        if (isSave) dt.Rows.Add(dataRow);
                    }
                    catch
                    {
                        Errors.Add(i + 1);
                    }

                    if (i % 500 == 0 || i >= sheet.LastRowNum)
                    {
                        var newDt = dt.Copy();

                        dt = dt.Clone();

                        yield return newDt;
                    }
                }
            }
        }
    }
}